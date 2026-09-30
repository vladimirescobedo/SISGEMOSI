@echo off
chcp 65001 >nul
echo ============================================
echo   SISGEMOSI - Reparación del Sistema
echo ============================================
echo.
echo [1/4] Deteniendo procesos Python...
taskkill /F /IM python.exe 2>nul
taskkill /F /IM pythonw.exe 2>nul
timeout /t 2 /nobreak >nul
echo [2/4] Verificando integridad de la base de datos...
python -c "import sqlite3; conn=sqlite3.connect('sistema_prestamos.db'); print(conn.execute('PRAGMA integrity_check').fetchone()[0]); conn.close()"
echo [3/4] Aplicando checkpoint WAL...
python -c "import sqlite3; conn=sqlite3.connect('sistema_prestamos.db'); conn.execute('PRAGMA wal_checkpoint(TRUNCATE)'); print('WAL checkpoint completado'); conn.close()"
echo [4/4] Reiniciando el servidor...
start "SISGEMOSI" python app.py
echo.
echo [OK] Sistema reparado y reiniciado.
pause
