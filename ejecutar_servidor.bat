@echo off
title Sistema de Prestamos - Servidor
cd /d C:\sistema_prestamos
if not exist "venv\Scripts\python.exe" (
	echo No se encontro el entorno virtual de Python en C:\sistema_prestamos\venv
	exit /b 1
)

:iniciar_servidor
"venv\Scripts\python.exe" app.py
echo [%date% %time%] El servidor se detuvo. Reiniciando en 5 segundos...
timeout /t 5 /nobreak >nul
goto iniciar_servidor