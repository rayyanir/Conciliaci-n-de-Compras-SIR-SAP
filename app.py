"""
app.py — Aplicación web de Conciliación SAP vs SIR
Ejecutar: python app.py
Luego abrir: http://localhost:5000
"""
import os, json, smtplib, tempfile, traceback
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from flask import (Flask, render_template, request, redirect,
                   url_for, send_file, flash, jsonify, session)
from werkzeug.utils import secure_filename

from reconciliation import run_reconciliation, generate_excel, get_all_cost_centers

# ── Configuración ──────────────────────────────────────────────────────────────

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR  = os.path.join(BASE_DIR, 'uploads')
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = 'conciliacion-sap-sir-kfc-2026'

STATE_FILE = os.path.join(UPLOAD_DIR, 'state.json')

def _load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {
        'results':    None,
        'excel_path': None,
        'period':     None,
        'timestamp':  None
    }

def _save_state(state):
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        traceback.print_exc()

ALLOWED_EXT = {'xlsx', 'xls'}

def _allowed(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

# ── Configuración SMTP / CC ────────────────────────────────────────────────────

def load_env_file():
    env_path = os.path.join(BASE_DIR, '.env')
    if os.path.exists(env_path):
        try:
            with open(env_path, encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        k, v = line.split('=', 1)
                        key = k.strip()
                        if key.startswith('\ufeff'):
                            key = key[1:]
                        if key:
                            val = v.strip()
                            if len(val) >= 2 and val[0] == val[-1] and val[0] in ('"', "'"):
                                val = val[1:-1]
                            os.environ[key] = val
        except Exception:
            pass

def save_smtp_to_env(smtp_cfg):
    env_lines = [
        f"SMTP_HOST={smtp_cfg.get('host', '')}",
        f"SMTP_PORT={smtp_cfg.get('port', '587')}",
        f"SMTP_USER={smtp_cfg.get('user', '')}",
        f"SMTP_PASSWORD={smtp_cfg.get('password', '')}",
        f"SMTP_FROM_ADDR={smtp_cfg.get('from_addr', '')}",
        f"SMTP_USE_TLS={'true' if smtp_cfg.get('use_tls') else 'false'}",
        f"SMTP_USE_SSL={'true' if smtp_cfg.get('use_ssl') else 'false'}",
        f"SMTP_NO_AUTH={'true' if smtp_cfg.get('no_auth') else 'false'}"
    ]
    with open(os.path.join(BASE_DIR, '.env'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(env_lines) + '\n')

def load_config():
    load_env_file()
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding='utf-8') as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    cfg = loaded
        except Exception:
            pass
            
    # Ensure all required keys exist and are of correct types
    if 'smtp' not in cfg or not isinstance(cfg['smtp'], dict):
        cfg['smtp'] = {}
    if 'cost_centers' not in cfg or not isinstance(cfg['cost_centers'], dict):
        cfg['cost_centers'] = {}
    if 'fixed_cc' not in cfg or not isinstance(cfg['fixed_cc'], list):
        cfg['fixed_cc'] = []
        
    # Populate SMTP settings from environment variables with fallback
    cfg['smtp']['host'] = os.environ.get('SMTP_HOST', cfg['smtp'].get('host', ''))
    cfg['smtp']['port'] = os.environ.get('SMTP_PORT', cfg['smtp'].get('port', '587'))
    cfg['smtp']['user'] = os.environ.get('SMTP_USER', cfg['smtp'].get('user', ''))
    cfg['smtp']['password'] = os.environ.get('SMTP_PASSWORD', cfg['smtp'].get('password', ''))
    cfg['smtp']['from_addr'] = os.environ.get('SMTP_FROM_ADDR', cfg['smtp'].get('from_addr', ''))
    
    # Booleans
    cfg['smtp']['use_tls'] = os.environ.get('SMTP_USE_TLS', 'true' if cfg['smtp'].get('use_tls') else 'false').lower() in ('true', '1', 'yes')
    cfg['smtp']['use_ssl'] = os.environ.get('SMTP_USE_SSL', 'true' if cfg['smtp'].get('use_ssl') else 'false').lower() in ('true', '1', 'yes')
    cfg['smtp']['no_auth'] = os.environ.get('SMTP_NO_AUTH', 'true' if cfg['smtp'].get('no_auth') else 'false').lower() in ('true', '1', 'yes')
    
    return cfg

def save_config(cfg):
    smtp_cfg = cfg.get('smtp', {})
    if smtp_cfg:
        try:
            save_smtp_to_env(smtp_cfg)
        except Exception as e:
            raise RuntimeError(f"Error al escribir en el archivo .env: {e}. Verifique los permisos de escritura del servidor.")
        
    clean_cfg = {k: v for k, v in cfg.items() if k != 'smtp'}
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(clean_cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        raise RuntimeError(f"Error al escribir en el archivo config.json: {e}. Verifique los permisos de escritura del servidor.")

# ── Envío de correo ────────────────────────────────────────────────────────────

def _send_email(smtp_cfg, to_addr, subject, html_body, cc_addrs=None):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From']    = smtp_cfg.get('from_addr') or smtp_cfg.get('user', '')
    msg['To']      = to_addr
    if cc_addrs:
        msg['Cc'] = ', '.join(cc_addrs)
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    all_recipients = [to_addr] + (cc_addrs or [])

    host     = smtp_cfg['host']
    port     = int(smtp_cfg.get('port', 25))
    user     = smtp_cfg.get('user', '')
    pwd      = smtp_cfg.get('password', '')
    no_auth  = smtp_cfg.get('no_auth', False)   # relay interno sin autenticación
    use_ssl  = smtp_cfg.get('use_ssl', False) or port == 465
    use_tls  = smtp_cfg.get('use_tls', False) and not use_ssl and not no_auth

    if use_ssl:
        with smtplib.SMTP_SSL(host, port, timeout=15) as srv:
            if not no_auth and user and pwd:
                srv.login(user, pwd)
            srv.sendmail(msg['From'], all_recipients, msg.as_string())
    else:
        with smtplib.SMTP(host, port, timeout=15) as srv:
            srv.ehlo()
            if use_tls:
                srv.starttls()
                srv.ehlo()
            if not no_auth and user and pwd:
                try:
                    srv.login(user, pwd)
                except smtplib.SMTPNotSupportedError:
                    pass   # servidor no requiere auth — continuar igual
            srv.sendmail(msg['From'], all_recipients, msg.as_string())


def _build_email_html(cc, cc_data, period):
    """Genera el cuerpo HTML del correo para un centro de costo."""
    ok       = cc_data['ok']
    diff     = cc_data['diff']
    only_sap = cc_data['only_sap']
    only_sir = cc_data['only_sir']
    no_fac   = cc_data['no_fac']

    def fmt(v):
        try: return f'${float(v):,.2f}'
        except: return str(v) if v else '-'

    total_issues = len(diff) + len(only_sap) + len(only_sir)
    status_color = '#2d7a2d' if total_issues == 0 else '#c07000'
    status_text  = 'Sin diferencias pendientes' if total_issues == 0 else f'{total_issues} diferencia(s) pendiente(s)'

    html = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<style>
  body{{font-family:Arial,sans-serif;color:#333;margin:0;padding:0;background:#f4f4f4}}
  .wrap{{max-width:700px;margin:20px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)}}
  .header{{background:#1F3864;color:#fff;padding:24px 28px}}
  .header h1{{margin:0;font-size:20px}}
  .header p{{margin:6px 0 0;font-size:13px;opacity:.85}}
  .body{{padding:24px 28px}}
  .status{{display:inline-block;background:{status_color};color:#fff;padding:6px 14px;border-radius:20px;font-size:13px;font-weight:bold;margin-bottom:20px}}
  .summary{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:24px}}
  .card{{flex:1;min-width:110px;background:#f0f4fa;border-radius:6px;padding:12px;text-align:center}}
  .card .num{{font-size:26px;font-weight:bold;color:#1F3864}}
  .card .lbl{{font-size:11px;color:#666;margin-top:4px}}
  .card.warn .num{{color:#c07000}}
  .card.err  .num{{color:#c00}}
  .card.ok   .num{{color:#2d7a2d}}
  h3{{color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-top:24px;font-size:14px}}
  table{{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}}
  th{{background:#1F3864;color:#fff;padding:7px 8px;text-align:left}}
  td{{padding:6px 8px;border-bottom:1px solid #e8e8e8}}
  tr:nth-child(even){{background:#f8f8f8}}
  .diff-val{{color:#c07000;font-weight:bold}}
  .footer{{background:#f4f4f4;padding:14px 28px;font-size:11px;color:#999;text-align:center}}
</style></head><body>
<div class="wrap">
  <div class="header">
    <h1>Conciliación SAP-SIR — Centro de Costo {cc}</h1>
    <p>Período: {period} &nbsp;|&nbsp; Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
  </div>
  <div class="body">
    <span class="status">{status_text}</span>
    <div class="summary">
      <div class="card ok"><div class="num">{len(ok)}</div><div class="lbl">Coincidencias OK</div></div>
      <div class="card {'warn' if diff else ''}"><div class="num">{len(diff)}</div><div class="lbl">Diferencias de monto</div></div>
      <div class="card {'err' if only_sap else ''}"><div class="num">{len(only_sap)}</div><div class="lbl">Solo en SAP</div></div>
      <div class="card {'err' if only_sir else ''}"><div class="num">{len(only_sir)}</div><div class="lbl">Solo en SIR</div></div>
    </div>
"""

    if diff:
        html += """<h3>⚠ Diferencias de Monto</h3>
<table><tr><th>Proveedor</th><th>Ref SAP</th><th>Factura SIR</th>
<th>Monto SAP</th><th>Monto SIR</th><th>Diferencia</th></tr>"""
        for r in sorted(diff, key=lambda x: abs(x.get('diferencia') or 0), reverse=True):
            html += (f"<tr><td>{r['vendor']}</td><td>{r['sap_ref']}</td>"
                     f"<td>{r['sir_factura']}</td><td>{fmt(r['sap_monto'])}</td>"
                     f"<td>{fmt(r['sir_total'])}</td>"
                     f"<td class='diff-val'>{fmt(r['diferencia'])}</td></tr>")
        html += '</table>'

    if only_sap:
        html += """<h3>🔵 Facturas en SAP no encontradas en SIR</h3>
<table><tr><th>Proveedor</th><th>Ref SAP</th><th>Últ. 5 dígitos</th><th>Monto SAP</th><th>Descripción</th></tr>"""
        for r in sorted(only_sap, key=lambda x: x.get('ref', '')):
            html += (f"<tr><td>{r['vendor']}</td><td>{r['ref']}</td>"
                     f"<td>{r['inv5']}</td><td>{fmt(r['sap_monto'])}</td>"
                     f"<td>{r.get('texto_cab','')}</td></tr>")
        html += '</table>'

    if only_sir:
        html += """<h3>🔴 Facturas en SIR no encontradas en SAP</h3>
<table><tr><th>Proveedor</th><th>Factura SIR</th><th>Monto SIR</th><th>Fecha</th></tr>"""
        for r in sorted(only_sir, key=lambda x: x.get('sir_factura', '')):
            html += (f"<tr><td>{r['vendor']}</td><td>{r['sir_factura']}</td>"
                     f"<td>{fmt(r['sir_total'])}</td><td>{r.get('sir_fecha','')}</td></tr>")
        html += '</table>'

    if total_issues == 0:
        html += '<p style="color:#2d7a2d;font-size:14px;margin-top:20px;">✔ Todas las facturas de este centro de costo están correctamente conciliadas.</p>'

    html += """
  </div>
  <div class="footer">
    Generado automáticamente por el Sistema de Conciliación SAP-SIR &nbsp;|&nbsp;
    Este correo es informativo — no responder directamente.
  </div>
</div></body></html>"""
    return html


# ── Rutas ──────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET'])
def index():
    cfg = load_config()
    cc_emails = cfg.get('cost_centers', {})
    state = _load_state()
    return render_template('index.html',
                           results=state['results'],
                           period=state['period'],
                           timestamp=state['timestamp'],
                           cc_emails=cc_emails)


@app.route('/compare', methods=['POST'])
def compare():
    sap_f    = request.files.get('sap')
    pepsi_f  = request.files.get('pepsi')
    larkin_f = request.files.get('larkin')
    period   = request.form.get('period', 'Mayo 2026').strip() or 'Mayo 2026'

    if not sap_f or not pepsi_f or not larkin_f:
        flash('Debes cargar los tres archivos (SAP, SIR PEPSI y SIR LARKIN).', 'danger')
        return redirect(url_for('index'))

    for f, label in [(sap_f, 'SAP'), (pepsi_f, 'SIR PEPSI'), (larkin_f, 'SIR LARKIN')]:
        if not _allowed(f.filename):
            flash(f'El archivo {label} debe ser .xlsx o .xls', 'danger')
            return redirect(url_for('index'))

    sap_path    = os.path.join(UPLOAD_DIR, 'sap_upload.xlsx')
    pepsi_path  = os.path.join(UPLOAD_DIR, 'pepsi_upload.xlsx')
    larkin_path = os.path.join(UPLOAD_DIR, 'larkin_upload.xls')

    sap_f.save(sap_path)
    pepsi_f.save(pepsi_path)
    larkin_f.save(larkin_path)

    try:
        results = run_reconciliation(sap_path, pepsi_path, larkin_path)
    except Exception as e:
        flash(f'Error al procesar los archivos: {e}', 'danger')
        traceback.print_exc()
        return redirect(url_for('index'))

    period_slug = period.replace(' ', '_').replace('/', '-')
    excel_path  = os.path.join(UPLOAD_DIR, f'Conciliacion_{period_slug}.xlsx')
    try:
        generate_excel(results, excel_path, period)
    except Exception as e:
        flash(f'Error al generar el Excel: {e}', 'danger')
        traceback.print_exc()
        return redirect(url_for('index'))

    state = {
        'results':    results,
        'excel_path': excel_path,
        'period':     period,
        'timestamp':  datetime.now().strftime('%d/%m/%Y %H:%M')
    }
    _save_state(state)

    flash(f'Conciliación completada correctamente — {period}', 'success')
    return redirect(url_for('index'))


@app.route('/download')
def download():
    state = _load_state()
    if not state['excel_path'] or not os.path.exists(state['excel_path']):
        flash('No hay reporte generado. Ejecuta la conciliación primero.', 'warning')
        return redirect(url_for('index'))
    period_slug = (state['period'] or 'reporte').replace(' ', '_')
    return send_file(state['excel_path'],
                     as_attachment=True,
                     download_name=f'Conciliacion_SAP_SIR_{period_slug}.xlsx')


@app.route('/config', methods=['GET', 'POST'])
def config():
    cfg = load_config()

    if request.method == 'POST':
        action = request.form.get('action', 'save')

        if action == 'save':
            cfg['smtp'] = {
                'host':      request.form.get('smtp_host', '').strip(),
                'port':      request.form.get('smtp_port', '25').strip(),
                'user':      request.form.get('smtp_user', '').strip(),
                'password':  request.form.get('smtp_password', ''),
                'from_addr': request.form.get('smtp_from', '').strip(),
                'use_tls':   request.form.get('use_tls') == 'on',
                'use_ssl':   request.form.get('use_ssl') == 'on',
                'no_auth':   request.form.get('no_auth') == 'on',
            }
            # Correos en copia fija
            fixed_raw = request.form.get('fixed_cc', '')
            cfg['fixed_cc'] = [e.strip() for e in fixed_raw.replace(';', ',').split(',') if e.strip()]
            ccs    = request.form.getlist('cc_code')
            emails = request.form.getlist('cc_email')
            names  = request.form.getlist('cc_name')
            cfg['cost_centers'] = {}
            for cc, email, name in zip(ccs, emails, names):
                cc = cc.strip(); email = email.strip()
                if cc:
                    cfg['cost_centers'][cc] = {'email': email, 'name': name.strip()}
            try:
                save_config(cfg)
                flash('Configuración guardada exitosamente.', 'success')
            except Exception as e:
                flash(str(e), 'danger')
            return redirect(url_for('config'))

    # Auto-poblar CCs desde últimos resultados
    state = _load_state()
    all_ccs = get_all_cost_centers(state['results']) if state['results'] else []
    existing = cfg.get('cost_centers', {})
    for cc in all_ccs:
        if cc not in existing:
            existing[cc] = {'email': '', 'name': ''}
    cfg['cost_centers'] = existing

    return render_template('config.html', cfg=cfg)


@app.route('/send-emails', methods=['POST'])
def send_emails():
    state = _load_state()
    if not state['results']:
        return jsonify({'ok': False, 'error': 'No hay resultados. Ejecuta la conciliación primero.'})

    cfg      = load_config()
    smtp_cfg = cfg.get('smtp', {})
    cc_cfg   = cfg.get('cost_centers', {})
    results  = state['results']
    period   = state['period'] or 'Sin especificar'

    if not smtp_cfg.get('host'):
        return jsonify({'ok': False, 'error': 'Configura el servidor SMTP en la página de Configuración.'})

    # Correos en copia fija (siempre incluidos en todos los envíos)
    fixed_cc = [e.strip() for e in cfg.get('fixed_cc', []) if e.strip()]

    sent, errors = [], []

    for cc, info in cc_cfg.items():
        email = info.get('email', '') if isinstance(info, dict) else str(info)
        if not email:
            continue

        cc_data = {
            'ok':       [r for r in results['matches_ok']   if r.get('cc') == cc],
            'diff':     [r for r in results['matches_diff'] if r.get('cc') == cc],
            'only_sap': [r for r in results['only_sap']     if r.get('cc') == cc],
            'only_sir': [r for r in results['only_sir']     if r.get('cc') == cc],
            'no_fac':   [r for r in results['sir_no_fac']   if r.get('cc') == cc],
        }

        name = info.get('name', cc) if isinstance(info, dict) else cc
        subject = f'Conciliación de Compras SAP-SIR {period} — Centro de Costo {cc}'
        if name and name != cc:
            subject = f'Conciliación de Compras SAP-SIR {period} — {cc} {name}'

        html_body = _build_email_html(cc, cc_data, period)

        try:
            _send_email(smtp_cfg, email, subject, html_body, cc_addrs=fixed_cc)
            sent.append({'cc': cc, 'email': email})
        except Exception as e:
            errors.append({'cc': cc, 'email': email, 'error': str(e)})

    return jsonify({'ok': True, 'sent': sent, 'errors': errors,
                    'total': len(sent), 'failed': len(errors)})


@app.route('/test-email', methods=['POST'])
def test_email():
    data     = request.get_json()
    smtp_cfg = data.get('smtp', {})
    to_addr  = data.get('to', '')
    if not to_addr:
        return jsonify({'ok': False, 'error': 'Ingresa un correo de destino.'})
    try:
        _send_email(smtp_cfg, to_addr,
                    'Prueba de conexión — Sistema Conciliación SAP-SIR',
                    '<p>Conexión SMTP configurada correctamente ✔</p>')
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/results-json')
def results_json():
    state = _load_state()
    if not state['results']:
        return jsonify({})
    return jsonify(state['results'])


# ── Arranque ───────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print('\n  Sistema de Conciliación SAP-SIR')
    print('  Abre tu navegador en: http://localhost:5000\n')
    app.run(debug=False, host='0.0.0.0', port=5000)
