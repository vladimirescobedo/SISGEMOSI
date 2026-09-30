@echo off
chcp 65001 >nul
echo ============================================
echo   SISGEMOSI - Crear Copia de Seguridad
echo ============================================
set FECHA=%date:~6,4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%%time:~6,2%
set FECHA=%FECHA: =0%
if not exist backups mkdir backups
copy /Y sistema_prestamos.db "backups\sistema_prestamos_%FECHA%.db"
copy /Y biblioteca_pae.db "backups\biblioteca_pae_%FECHA%.db"
echo.
echo [OK] Copia de seguridad creada en carpeta backups/
echo     - sistema_prestamos_%FECHA%.db
echo     - biblioteca_pae_%FECHA%.db
pause
