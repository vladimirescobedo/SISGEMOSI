@echo off
setlocal
cd /d %~dp0
if not exist "certificados\rootCA.pem" (
    echo No se encontro certificados\rootCA.pem
    pause
    exit /b 1
)
certutil -user -addstore Root "certificados\rootCA.pem"
if errorlevel 1 (
    echo No se pudo instalar la autoridad certificadora.
    pause
    exit /b 1
)
echo Certificado instalado correctamente para este usuario.
echo Cierre y vuelva a abrir el navegador antes de ingresar.
pause
