@echo off
chcp 65001 > nul
echo ============================================
echo  Mise a jour du Dashboard Bassin Ivoirien
echo ============================================
echo.
python "%~dp0update_dashboard.py"
echo.
if %errorlevel%==0 (
    echo [OK] Dashboard mis a jour avec succes!
    echo Ouvrez: D:\LCISSE\Downloads\IVOIRIAN_BASIN_DASHBOARD.html
) else (
    echo [ERREUR] La mise a jour a echoue.
)
echo.
pause
