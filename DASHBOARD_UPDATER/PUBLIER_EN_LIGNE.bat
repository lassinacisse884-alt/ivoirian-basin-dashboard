@echo off
chcp 65001 > nul
echo ================================================
echo   MISE A JOUR + PUBLICATION EN LIGNE
echo   Bassin Sedimentaire Ivoirien - Dashboard
echo ================================================
echo.

:: --- 1. Mise a jour du HTML depuis Excel ---
echo [ETAPE 1/3] Mise a jour du dashboard depuis Excel...
python "%~dp0update_dashboard.py"
if %errorlevel% neq 0 (
    echo [ERREUR] La mise a jour du dashboard a echoue.
    pause
    exit /b 1
)
echo [OK] Dashboard HTML mis a jour.
echo.

:: --- 2. Git add + commit ---
echo [ETAPE 2/3] Preparation de la publication...
cd /d "D:\LCISSE\Desktop\IVOIRIAN SEDIMENTARY BASIN OIL AND GAS\1-GENERAL DATABASE"

git add IVOIRIAN_BASIN_DASHBOARD.html

:: Commit avec date/heure
set DATEHEURE=%date% %time:~0,5%
git commit -m "Dashboard mis a jour le %DATEHEURE%"
if %errorlevel% neq 0 (
    echo [INFO] Aucune modification detectee, pas de commit necessaire.
    echo.
    pause
    exit /b 0
)
echo.

:: --- 3. Push vers GitHub ---
echo [ETAPE 3/3] Publication en ligne...
git push origin main
if %errorlevel% neq 0 (
    echo [ERREUR] La publication a echoue. Verifiez votre connexion internet.
    echo          Si c'est la premiere fois, suivez les instructions de configuration.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   [SUCCES] Dashboard publie en ligne !
echo   Vos collegues peuvent y acceder via :
echo   https://lassinacisse884-alt.github.io/ivoirian-basin-dashboard/
echo ================================================
echo.
pause
