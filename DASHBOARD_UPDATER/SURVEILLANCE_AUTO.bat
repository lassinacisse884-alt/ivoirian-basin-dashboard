@echo off
chcp 65001 > nul
echo ============================================
echo  Surveillance automatique du Dashboard
echo ============================================
echo.
echo Le dashboard se mettra a jour automatiquement
echo des que le fichier Excel est modifie.
echo.
echo Appuyez sur Ctrl+C pour arreter.
echo.
python "%~dp0update_dashboard.py" --watch
pause
