"""
reconciliation.py — Lógica de conciliación SAP vs SIR
Compatible con COMPRAS SAP.xlsx, COMPRAS SIR PEPSI.xlsx, COMPRAS SIR LARKIN.xls
"""
import os, struct, zlib, re
from xml.etree import ElementTree as ET
from html.parser import HTMLParser
from collections import defaultdict
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


# ── Parsers de archivos ────────────────────────────────────────────────────────

def _read_zip_entry(filepath, target):
    with open(filepath, 'rb') as f:
        data = f.read()
    pos = 0
    while True:
        idx = data.find(b'\x50\x4b\x03\x04', pos)
        if idx == -1:
            return None
        ver, flags, comp, mtime, mdate, crc, csz, usz, fnlen, exlen = \
            struct.unpack_from('<5H3I2H', data, idx + 4)
        name = data[idx + 30:idx + 30 + fnlen].decode('utf-8', 'replace')
        ds = idx + 30 + fnlen + exlen
        if name == target:
            return zlib.decompress(data[ds:ds + csz], -15) if comp == 8 else data[ds:ds + csz]
        pos = idx + 4


def _parse_xlsx_sheet(filepath, sheet_path='xl/worksheets/sheet1.xml'):
    with open(filepath, 'rb') as f:
        header = f.read(4)
    if header != b'\x50\x4b\x03\x04':
        filename = os.path.basename(filepath)
        if header.startswith(b'\xd0\xcf\x11\xe0'):
            raise ValueError(f"El archivo '{filename}' es un Excel antiguo (.xls). Se requiere un archivo Excel moderno (.xlsx).")
        raise ValueError(f"El archivo '{filename}' no es un archivo Excel (.xlsx) válido.")
    ss_xml = _read_zip_entry(filepath, 'xl/sharedStrings.xml')
    ss = []
    if ss_xml:
        root = ET.fromstring(ss_xml)
        for si in root.findall(f'{{{NS}}}si'):
            ss.append(''.join(t.text or '' for t in si.iter(f'{{{NS}}}t')))
    sheet_xml = _read_zip_entry(filepath, sheet_path)
    if not sheet_xml:
        return []
    root = ET.fromstring(sheet_xml)
    rows = []
    for row_el in root.iter(f'{{{NS}}}row'):
        row = {}
        for c in row_el:
            col = re.match(r'([A-Z]+)', c.get('r', ''))
            if not col:
                continue
            col = col.group(1)
            t = c.get('t', '')
            v_el = c.find(f'{{{NS}}}v')
            if v_el is None or v_el.text is None:
                row[col] = None
            elif t == 's':
                row[col] = ss[int(v_el.text)] if int(v_el.text) < len(ss) else ''
            else:
                row[col] = v_el.text
        rows.append(row)
    return rows


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.cur_row = []
        self.cur_cell = None

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.cur_row = []
        elif tag in ('td', 'th'):
            self.cur_cell = ''

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            if self.cur_cell is not None:
                self.cur_row.append(self.cur_cell.strip())
            self.cur_cell = None
        elif tag == 'tr':
            if any(c.strip() for c in self.cur_row):
                self.rows.append(self.cur_row[:])

    def handle_data(self, data):
        if self.cur_cell is not None:
            self.cur_cell += data


def _parse_html_table(filepath):
    with open(filepath, 'rb') as f:
        raw_data = f.read()
    
    # Detectar firmas binarias para dar un mensaje de error claro
    if raw_data.startswith(b'\x50\x4b\x03\x04'):
        raise ValueError(f"El archivo '{os.path.basename(filepath)}' parece ser un Excel moderno (.xlsx). Se requiere el reporte en formato HTML/xls.")
    if raw_data.startswith(b'\xd0\xcf\x11\xe0'):
        raise ValueError(f"El archivo '{os.path.basename(filepath)}' es un archivo de Excel binario (.xls real). El sistema requiere el reporte en formato HTML/xls.")
    
    # Intentar decodificar con distintas codificaciones
    content = None
    encodings = ['utf-8-sig', 'utf-16', 'cp1252', 'latin-1']
    for enc in encodings:
        try:
            content = raw_data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
            
    if content is None:
        raise ValueError(f"No se pudo decodificar el archivo '{os.path.basename(filepath)}'. Verifique la codificación del archivo.")
        
    p = _TableParser()
    p.feed(content)
    return p.rows


# ── Helpers ────────────────────────────────────────────────────────────────────

def _last5(v):
    if v is None:
        return None
    d = re.sub(r'[^0-9]', '', str(v))
    return d[-5:] if len(d) >= 5 else (d if d else None)


def _parse_num(v, european=False):
    if v is None or str(v).strip() in ('-', '', '–'):
        return None
    s = str(v).strip()
    if european:
        s = s.replace('.', '').replace(',', '.')
    try:
        return round(float(s), 2)
    except Exception:
        return None


def _excel_date(v):
    try:
        return (date(1899, 12, 30) + timedelta(days=int(float(v)))).strftime('%d/%m/%Y')
    except Exception:
        return str(v) if v else ''


# ── Carga de archivos ──────────────────────────────────────────────────────────

def _load_sap(path):
    rows = _parse_xlsx_sheet(path, 'xl/worksheets/sheet1.xml')
    sap_pepsi, sap_larkin = [], []
    for r in rows[1:]:
        if not any(r.values()):
            continue
        ref = str(r.get('D', '') or '').strip()
        cc_ben = str(r.get('L', '') or '').strip()
        cc = cc_ben[:-2] if cc_ben.endswith('VE') else cc_ben
        monto = _parse_num(r.get('I'))
        proveedor = str(r.get('K', '') or '').strip()
        texto_cab = str(r.get('E', '') or '').strip()
        inv5 = _last5(ref)
        rec = {'cc': cc, 'ref': ref, 'inv5': inv5, 'monto': monto,
               'proveedor': proveedor, 'texto_cab': texto_cab}
        if 'LARKIN' in proveedor.upper():
            sap_larkin.append(rec)
        else:
            sap_pepsi.append(rec)
    return sap_pepsi, sap_larkin


def _load_sir_pepsi(path):
    rows = _parse_xlsx_sheet(path, 'xl/worksheets/sheet2.xml')
    data_start = 0
    for i, r in enumerate(rows):
        if 'Centro de Costo' in str(r.get('A', '') or ''):
            data_start = i + 1
            break
    result = []
    last_cc = None
    for r in rows[data_start:]:
        cc_raw = r.get('A')
        if cc_raw and not str(cc_raw).startswith('='):
            last_cc = str(cc_raw).strip()
        cc = last_cc
        if cc is None:
            continue
        fac = str(r.get('D') or '').strip()
        fac5 = _last5(fac)
        total = _parse_num(r.get('M'))
        dev = _parse_num(r.get('L'))
        fecha = _excel_date(r.get('B'))
        cod = str(r.get('F') or '').strip()
        result.append({'cc': cc, 'factura': fac, 'inv5': fac5, 'total': total,
                       'devolucion': dev, 'fecha': fecha, 'cod': cod})
    return result


def _load_sir_larkin(path):
    html_rows = _parse_html_table(path)
    result = []
    last_cc = None
    for r in html_rows[5:]:
        if 'Sub Total' in ' '.join(r):
            continue
        if len(r) < 12:
            continue
        cc_raw = r[0].strip()
        if cc_raw and cc_raw != '-':
            last_cc = cc_raw
        cc = last_cc
        if cc is None:
            continue
        fac = r[3].strip()
        fac5 = _last5(fac)
        total = _parse_num(r[11], european=True)
        dev = _parse_num(r[10], european=True)
        fecha = r[1].strip()
        cod = r[4].strip() if len(r) > 4 else ''
        cod_dev = r[5].strip() if len(r) > 5 else ''
        result.append({'cc': cc, 'factura': fac, 'inv5': fac5, 'total': total,
                       'devolucion': dev, 'fecha': fecha, 'cod': cod, 'cod_dev': cod_dev})
    return result


# ── Motor de comparación ───────────────────────────────────────────────────────

def _compare(sap_list, sir_list, tolerance, vendor_label):
    sir_lu = defaultdict(list)
    sir_no_inv = []
    for r in sir_list:
        if r['inv5']:
            sir_lu[(r['cc'], r['inv5'])].append(r)
        else:
            sir_no_inv.append(r)

    sap_lu = defaultdict(list)
    for r in sap_list:
        if r['inv5'] and r['cc']:
            sap_lu[(r['cc'], r['inv5'])].append(r)

    sir_keys = set(sir_lu)
    sap_keys = set(sap_lu)
    matches_ok, matches_diff, only_sir, only_sap, sir_no_fac = [], [], [], [], []

    for key in sir_keys & sap_keys:
        for sr in sir_lu[key]:
            for sp in sap_lu[key]:
                sa, spa = sr['total'], sp['monto']
                diff = abs(sa - spa) if (sa is not None and spa is not None) else None
                row = {'vendor': vendor_label, 'cc': key[0], 'inv5': key[1],
                       'sir_factura': sr['factura'], 'sap_ref': sp['ref'],
                       'sir_total': sa, 'sap_monto': spa,
                       'diferencia': round(spa - sa, 2) if diff is not None else None,
                       'sir_fecha': sr['fecha'], 'sir_cod': sr.get('cod', ''),
                       'proveedor': sp['proveedor'], 'tolerancia': tolerance}
                (matches_ok if (diff is not None and diff <= tolerance) else matches_diff).append(row)

    for key in sir_keys - sap_keys:
        for sr in sir_lu[key]:
            only_sir.append({'vendor': vendor_label, 'cc': key[0], 'inv5': key[1],
                             'sir_factura': sr['factura'], 'sir_total': sr['total'],
                             'sir_fecha': sr['fecha'], 'sir_cod': sr.get('cod', '')})

    matched_no_inv = set()
    for key in sap_keys - sir_keys:
        for sp in sap_lu[key]:
            if sp['monto'] is not None and sp['monto'] < 0:
                continue
            paired = False
            for idx, sr in enumerate(sir_no_inv):
                if idx in matched_no_inv:
                    continue
                if sr['cc'] != sp['cc']:
                    continue
                sa, spa = sr['total'], sp['monto']
                diff = abs(sa - spa) if (sa is not None and spa is not None) else None
                if diff is not None and diff <= tolerance:
                    matched_no_inv.add(idx)
                    matches_ok.append({
                        'vendor': vendor_label, 'cc': sp['cc'], 'inv5': sp['inv5'],
                        'sir_factura': '(sin N° SIR)', 'sap_ref': sp['ref'],
                        'sir_total': sa, 'sap_monto': spa,
                        'diferencia': round(spa - sa, 2),
                        'sir_fecha': sr.get('fecha', ''), 'sir_cod': sr.get('cod', ''),
                        'proveedor': sp['proveedor'], 'tolerancia': tolerance
                    })
                    paired = True
                    break
            if not paired:
                only_sap.append({'vendor': vendor_label, 'cc': sp['cc'], 'inv5': sp['inv5'],
                                 'ref': sp['ref'], 'sap_monto': sp['monto'],
                                 'proveedor': sp['proveedor'], 'texto_cab': sp['texto_cab']})

    for idx, r in enumerate(sir_no_inv):
        if idx not in matched_no_inv:
            nf = r.copy()
            nf['vendor'] = vendor_label
            sir_no_fac.append(nf)

    return matches_ok, matches_diff, only_sir, only_sap, sir_no_fac


# ── API pública ────────────────────────────────────────────────────────────────

def run_reconciliation(sap_path, pepsi_path, larkin_path):
    """Ejecuta la conciliación completa. Devuelve dict con todos los resultados."""
    sap_pepsi, sap_larkin = _load_sap(sap_path)
    sir_pepsi = _load_sir_pepsi(pepsi_path)
    sir_larkin = _load_sir_larkin(larkin_path)

    p_ok, p_diff, p_sir, p_sap, p_nofac = _compare(sap_pepsi, sir_pepsi, 1.0, 'PEPSI')
    l_ok, l_diff, l_sir, l_sap, l_nofac = _compare(sap_larkin, sir_larkin, 5.0, 'LARKIN')

    return {
        'summary': {
            'pepsi': {
                'ok': len(p_ok), 'diff': len(p_diff),
                'only_sir': len(p_sir), 'only_sap': len(p_sap), 'no_fac': len(p_nofac)
            },
            'larkin': {
                'ok': len(l_ok), 'diff': len(l_diff),
                'only_sir': len(l_sir), 'only_sap': len(l_sap), 'no_fac': len(l_nofac)
            }
        },
        'matches_ok':  p_ok  + l_ok,
        'matches_diff': p_diff + l_diff,
        'only_sir':    p_sir  + l_sir,
        'only_sap':    p_sap  + l_sap,
        'sir_no_fac':  p_nofac + l_nofac
    }


def get_all_cost_centers(results):
    """Devuelve lista ordenada de todos los centros de costo presentes en los resultados."""
    ccs = set()
    for lst in [results['matches_ok'], results['matches_diff'],
                results['only_sir'], results['only_sap'], results['sir_no_fac']]:
        for r in lst:
            if r.get('cc'):
                ccs.add(r['cc'])
    return sorted(ccs)


# ── Generación del reporte Excel ───────────────────────────────────────────────

def generate_excel(results, output_path, period='Mayo 2026'):
    """Genera el reporte Excel de conciliación."""
    thin = Side(style='thin', color='BFBFBF')
    B = Border(left=thin, right=thin, top=thin, bottom=thin)

    def mk_fill(hex_): return PatternFill('solid', start_color=hex_)
    def mk_font(**kw): return Font(name='Arial', **kw)

    H_FILL = mk_fill('1F3864'); H_FONT = mk_font(bold=True, color='FFFFFF', size=11)
    H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    OK_F = mk_fill('E2EFDA'); WARN_F = mk_fill('FFF2CC')
    ERR_F = mk_fill('FCE4D6'); BLUE_F = mk_fill('DAEEF3')
    DIFF_F = mk_fill('FFD966'); TOT_F = mk_fill('2F5496')
    P_HDR = mk_fill('1F5C99'); L_HDR = mk_fill('7B2D8B')
    P_ROW = mk_fill('E9F0FB'); L_ROW = mk_fill('F5E6FA')
    P_ROW2 = mk_fill('D0E4F7'); L_ROW2 = mk_fill('EED5FA')
    GRAY_F = mk_fill('F2F2F2')

    def hdr(ws, row, cols, fill=None):
        for c in range(1, cols + 1):
            cell = ws.cell(row, c)
            cell.font = H_FONT; cell.fill = fill or H_FILL
            cell.alignment = H_ALIGN; cell.border = B

    def brow(ws, row_i, cols, fill):
        for c in range(1, cols + 1):
            cell = ws.cell(row_i, c)
            cell.font = mk_font(size=10); cell.border = B
            if fill: cell.fill = fill
            cell.alignment = Alignment(
                horizontal='right' if isinstance(cell.value, (int, float)) else 'left')

    def tot_row(ws, row_i, cols, sum_cols):
        ws.cell(row_i, 1, 'TOTAL')
        for c in sum_cols:
            ws.cell(row_i, c,
                    f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}{row_i - 1})')
        for c in range(1, cols + 1):
            ws.cell(row_i, c).font = mk_font(bold=True, color='FFFFFF', size=10)
            ws.cell(row_i, c).fill = TOT_F; ws.cell(row_i, c).border = B

    def title(ws, text, cols, hex_):
        ws.row_dimensions[1].height = 36
        ws.merge_cells(f'A1:{get_column_letter(cols)}1')
        ws['A1'] = text
        ws['A1'].font = mk_font(bold=True, size=13, color='FFFFFF')
        ws['A1'].fill = mk_fill(hex_)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    def col_w(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def section_hdr(ws, row_i, text, cols, fill):
        ws.merge_cells(f'A{row_i}:{get_column_letter(cols)}{row_i}')
        ws.cell(row_i, 1, text)
        ws.cell(row_i, 1).font = mk_font(bold=True, color='FFFFFF', size=11)
        ws.cell(row_i, 1).fill = fill
        ws.cell(row_i, 1).alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[row_i].height = 20

    p_ok   = [r for r in results['matches_ok']   if r['vendor'] == 'PEPSI']
    l_ok   = [r for r in results['matches_ok']   if r['vendor'] == 'LARKIN']
    p_diff = [r for r in results['matches_diff'] if r['vendor'] == 'PEPSI']
    l_diff = [r for r in results['matches_diff'] if r['vendor'] == 'LARKIN']
    p_sir  = [r for r in results['only_sir']     if r['vendor'] == 'PEPSI']
    l_sir  = [r for r in results['only_sir']     if r['vendor'] == 'LARKIN']
    p_sap  = [r for r in results['only_sap']     if r['vendor'] == 'PEPSI']
    l_sap  = [r for r in results['only_sap']     if r['vendor'] == 'LARKIN']
    p_nofac = [r for r in results['sir_no_fac']  if r['vendor'] == 'PEPSI']
    l_nofac = [r for r in results['sir_no_fac']  if r['vendor'] == 'LARKIN']

    wb = Workbook()

    # ── RESUMEN ──
    ws = wb.active; ws.title = 'RESUMEN'
    ws.row_dimensions[1].height = 44
    ws.merge_cells('A1:E1')
    ws['A1'] = f'CONCILIACIÓN SAP vs SIR — PEPSI & LARKIN — {period.upper()}'
    ws['A1'].font = mk_font(bold=True, size=14, color='FFFFFF')
    ws['A1'].fill = mk_fill('1F3864')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.append(['Proveedor', 'Resultado', 'Descripción', 'Registros', 'Tolerancia ($)'])
    hdr(ws, 2, 5)
    ws.row_dimensions[2].height = 22

    def s_row(ws, vendor, lbl, desc, cnt, tol, fill, vfill):
        r = ws.max_row + 1
        ws.cell(r, 1, vendor); ws.cell(r, 2, lbl); ws.cell(r, 3, desc)
        ws.cell(r, 4, cnt); ws.cell(r, 5, tol)
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.font = mk_font(size=11, bold=(c == 4))
            cell.fill = vfill if c == 1 else fill; cell.border = B
            cell.alignment = Alignment(
                horizontal='center' if c in (1, 4, 5) else 'left', vertical='center')
        ws.row_dimensions[r].height = 20

    for lbl, desc, cnt, fill in [
        ('✔ OK', 'Diferencia ≤ $1', len(p_ok), OK_F),
        ('⚠ Diferencia', 'Diferencia > $1', len(p_diff), WARN_F),
        ('Solo en SIR', 'En SIR, no en SAP', len(p_sir), ERR_F),
        ('Solo en SAP', 'En SAP, no en SIR (monto ≥ $0)', len(p_sap), BLUE_F)]:
        s_row(ws, 'PEPSI', lbl, desc, cnt, '$1.00', fill, P_HDR)
    for lbl, desc, cnt, fill in [
        ('✔ OK', 'Diferencia ≤ $5', len(l_ok), OK_F),
        ('⚠ Diferencia', 'Diferencia > $5', len(l_diff), WARN_F),
        ('Solo en SIR', 'En SIR, no en SAP', len(l_sir), ERR_F),
        ('Solo en SAP', 'En SAP, no en SIR (monto ≥ $0)', len(l_sap), BLUE_F)]:
        s_row(ws, 'LARKIN', lbl, desc, cnt, '$5.00', fill, L_HDR)
    ws.cell(ws.max_row + 2, 1, 'Matching: últimos 5 dígitos del N° de factura | Montos en USD')
    ws.cell(ws.max_row, 1).font = mk_font(italic=True, size=9, color='666666')
    col_w(ws, [12, 20, 46, 12, 14])

    # ── DIFERENCIAS DE MONTO ──
    ws2 = wb.create_sheet('Diferencias de Monto')
    NC = 8
    title(ws2, f'FACTURAS CON DIFERENCIA DE MONTO POR ENCIMA DE LA TOLERANCIA — {period.upper()}', NC, 'BF8F00')
    ws2.append(['Proveedor', 'Centro de Costo', 'Últ. 5 Fac.', 'Factura SIR', 'Fecha SIR',
                'Monto SIR ($)', 'Monto SAP ($)', 'Diferencia ($)'])
    hdr(ws2, 2, NC)

    for vendor, rows, ra, rb, hfill in [
        ('PEPSI', p_diff, P_ROW, P_ROW2, P_HDR),
        ('LARKIN', l_diff, L_ROW, L_ROW2, L_HDR)
    ]:
        if rows:
            section_hdr(ws2, ws2.max_row + 1,
                        f'── {vendor}  (tolerancia {"$1" if vendor == "PEPSI" else "$5"}) ──', NC, hfill)
            for i, r in enumerate(sorted(rows, key=lambda x: (x['cc'], x['inv5'])), 1):
                ws2.append([r['vendor'], r['cc'], r['inv5'], r['sir_factura'], r['sir_fecha'],
                            r['sir_total'], r['sap_monto'], r['diferencia']])
                brow(ws2, ws2.max_row, NC, ra if i % 2 == 0 else rb)
                ws2.cell(ws2.max_row, 8).fill = DIFF_F
                ws2.cell(ws2.max_row, 8).font = mk_font(size=10, bold=True)
    tot_row(ws2, ws2.max_row + 1, NC, [6, 7, 8])
    col_w(ws2, [10, 16, 12, 16, 12, 14, 14, 14])

    # ── SOLO EN SIR ──
    ws3 = wb.create_sheet('Solo en SIR')
    NC3 = 6
    title(ws3, f'FACTURAS EN SIR QUE NO ESTÁN EN SAP — {period.upper()}', NC3, 'C00000')
    ws3.append(['Proveedor', 'Centro de Costo', 'Últ. 5 Fac.', 'Factura SIR', 'Fecha SIR', 'Monto SIR ($)'])
    hdr(ws3, 2, NC3)
    for vendor, rows, ra, rb, hfill in [
        ('PEPSI', p_sir, ERR_F, mk_fill('FCF0ED'), P_HDR),
        ('LARKIN', l_sir, L_ROW, L_ROW2, L_HDR)
    ]:
        if rows:
            section_hdr(ws3, ws3.max_row + 1, f'── {vendor} ──', NC3, hfill)
            for i, r in enumerate(sorted(rows, key=lambda x: (x['cc'], str(x.get('sir_factura', '')))), 1):
                ws3.append([r['vendor'], r['cc'], r['inv5'], r.get('sir_factura', ''),
                            r.get('sir_fecha', ''), r.get('sir_total')])
                brow(ws3, ws3.max_row, NC3, ra if i % 2 == 0 else rb)
    tot_row(ws3, ws3.max_row + 1, NC3, [6])
    col_w(ws3, [10, 16, 12, 18, 12, 14])

    # ── SOLO EN SAP ──
    ws4 = wb.create_sheet('Solo en SAP')
    NC4 = 6
    title(ws4, f'FACTURAS EN SAP QUE NO ESTÁN EN SIR — {period.upper()}', NC4, '1F497D')
    ws4.append(['Proveedor', 'Centro de Costo', 'Ref SAP', 'Últ. 5 Fac.', 'Monto SAP ($)', 'Descripción'])
    hdr(ws4, 2, NC4)
    for vendor, rows, ra, rb, hfill in [
        ('PEPSI', p_sap, BLUE_F, mk_fill('E4EFF7'), P_HDR),
        ('LARKIN', l_sap, L_ROW, L_ROW2, L_HDR)
    ]:
        if rows:
            section_hdr(ws4, ws4.max_row + 1, f'── {vendor} ──', NC4, hfill)
            for i, r in enumerate(sorted(rows, key=lambda x: (x['cc'], str(x.get('ref', '')))), 1):
                ws4.append([r['vendor'], r['cc'], r['ref'], r['inv5'],
                            r['sap_monto'], r.get('texto_cab', '')])
                brow(ws4, ws4.max_row, NC4, ra if i % 2 == 0 else rb)
    tot_row(ws4, ws4.max_row + 1, NC4, [5])
    col_w(ws4, [10, 16, 22, 12, 14, 35])

    # ── SIR SIN FACTURA ──
    ws5 = wb.create_sheet('SIR sin Factura')
    NC5 = 6
    title(ws5, f'REGISTROS SIR SIN NÚMERO DE FACTURA — {period.upper()}', NC5, '595959')
    ws5.append(['Proveedor', 'Centro de Costo', 'Fecha', 'Código Compras', 'Código Dev.', 'Monto ($)'])
    hdr(ws5, 2, NC5)
    all_nofac = p_nofac + l_nofac
    for i, r in enumerate(sorted(all_nofac, key=lambda x: (x['cc'], x.get('fecha', ''))), 3):
        monto = r.get('devolucion') if r.get('devolucion') else r.get('total')
        ws5.append([r['vendor'], r['cc'], r.get('fecha', ''), r.get('cod', ''),
                    r.get('cod_dev', ''), monto])
        brow(ws5, ws5.max_row, NC5, GRAY_F if i % 2 == 0 else mk_fill('FFFFFF'))
    col_w(ws5, [10, 16, 12, 22, 22, 14])

    wb.save(output_path)
