@echo off
title Conciliacion SAP-SIR
echo.
echo  ============================================
echo   Sistema de Conciliacion SAP-SIR
echo  ============================================
echo.

cd /d "%~dp0"

echo  Verificando dependencias...
pip install -r requirements.txt --quiet

echo.
echo  Abriendo puerto 5000 en el firewall de Windows...
netsh advfirewall firewall delete rule name="ConciliacionSAPSIR" >nul 2>&1
netsh advfirewall firewall add rule name="ConciliacionSAPSIR" dir=in action=allow protocol=TCP localport=5000 >nul 2>&1

echo.
echo  ============================================
echo   Servidor iniciado
echo  ============================================
echo.
echo  Desde ESTA maquina:
echo    http://localhost:5000
echo.
echo  Desde OTROS computadores en la red:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" ^| findstr /v "169.254"') do (
    set IP=%%a
    goto :show_ip
)
:show_ip
set IP=%IP: =%
echo    http://%IP%:5000
echo.
echo  Comparte esa direccion con tus usuarios.
echo  Mantener esta ventana abierta mientras se usa la aplicacion.
echo  ============================================
echo.

start "" "http://localhost:5000"
python app.py

echo.
echo  Servidor detenido.
pause
