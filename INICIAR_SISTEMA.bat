@echo off
chcp 65001 >nul
title SISGEMOSI - UT Apurímac PAE
echo ==========================================================
echo    SISTEMA DE GESTIÓN DE MÓDULOS DE SOPORTE INFORMÁTICO
echo                 SISGEMOSI - UT APURÍMAC
echo ==========================================================
echo.

cd /d "%~dp0"

echo [1/3] Verificando entorno virtual Python...
if exist "venv\Scripts\python.exe" (
    set "PYTHON_EXE=venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

echo [2/3] Verificando base de datos y dependencias...
"%PYTHON_EXE%" -c "from database import init_db, init_biblioteca_db; init_db(); init_biblioteca_db(); print('[OK] Bases de datos verificadas.')"

echo [3/3] Iniciando servidor SISGEMOSI en https://localhost:5000...
echo.
echo Presione Ctrl+C para detener el servidor en cualquier momento.
echo ==========================================================
echo.

start "" https://localhost:5000
"%PYTHON_EXE%" app.py
pause

