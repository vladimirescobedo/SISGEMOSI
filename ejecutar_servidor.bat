@echo off
title SISGEMOSI - Servidor HTTPS
cd /d C:\sistema_prestamos
if not exist "venv\Scripts\python.exe" (
	echo No se encontro el entorno virtual de Python en C:\sistema_prestamos\venv
	exit /b 1
)
if not exist "certificados\sisgemosi-cert.pem" (
	echo No se encontro el certificado HTTPS en C:\sistema_prestamos\certificados
	exit /b 1
)
if not exist "certificados\sisgemosi-key.pem" (
	echo No se encontro la clave HTTPS en C:\sistema_prestamos\certificados
	exit /b 1
)

echo Cerrando instancias anteriores de SISGEMOSI...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$self = (Get-CimInstance Win32_Process -Filter ('ProcessId=' + $PID)).ParentProcessId; Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'cmd.exe' -and $_.CommandLine -match 'ejecutar_servidor[.]bat' -and $_.ProcessId -ne $self } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; $python = (Resolve-Path 'venv\Scripts\python.exe').Path; Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $python -and $_.CommandLine -match 'app[.]py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>&1
timeout /t 1 /nobreak >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$rule = Get-NetFirewallRule -DisplayName 'SISGEMOSI TCP 5000' -ErrorAction SilentlyContinue; if ($rule) { Set-NetFirewallRule -DisplayName 'SISGEMOSI TCP 5000' -Profile Any -Enabled True } else { New-NetFirewallRule -DisplayName 'SISGEMOSI TCP 5000' -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow -Profile Any }" >nul 2>&1

echo SISGEMOSI disponible en https://IP_DEL_EQUIPO:5000

:iniciar_servidor
"venv\Scripts\python.exe" app.py
echo [%date% %time%] El servidor se detuvo. Reiniciando en 5 segundos...
timeout /t 5 /nobreak >nul
goto iniciar_servidor