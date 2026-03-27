@echo off
setlocal

echo [CalCal] Installiere Abhaengigkeiten...
call npm install
if errorlevel 1 goto :err

echo [CalCal] Baue Portable EXE + Installer...
call npm run build-win-all
if errorlevel 1 goto :err

echo.
echo [CalCal] Fertig. Dateien liegen im Ordner dist\
echo  - CalCal-Portable-*.exe
echo  - CalCal-Setup-*.exe
echo.
pause
exit /b 0

:err
echo.
echo [CalCal] Fehler beim Build.
pause
exit /b 1
