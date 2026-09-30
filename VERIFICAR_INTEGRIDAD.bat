@echo off
chcp 65001 >nul
echo ============================================
echo   SISGEMOSI - Verificar Integridad BD
echo ============================================
echo.
echo Verificando sistema_prestamos.db...
python -c "import sqlite3; conn=sqlite3.connect('sistema_prestamos.db'); result=conn.execute('PRAGMA integrity_check').fetchone()[0]; print('Estado: ' + result); conn.close()"
echo.
echo Verificando biblioteca_pae.db...
python -c "import sqlite3; conn=sqlite3.connect('biblioteca_pae.db'); result=conn.execute('PRAGMA integrity_check').fetchone()[0]; print('Estado: ' + result); conn.close()"
echo.
echo Información de tablas:
python -c "import sqlite3; conn=sqlite3.connect('sistema_prestamos.db'); tables=conn.execute(\"SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%%'\").fetchall(); [print(f'  - {t[0]}: {conn.execute(f\"SELECT COUNT(*) FROM {t[0]}\").fetchone()[0]} registros') for t in tables]; conn.close()"
echo.
pause
