@echo off
chcp 65001 >nul
title AC Crew Pairings

echo ============================================
echo    Air Canada - Crew Pairings
echo ============================================
echo.

:: Check if data file exists
if not exist "%~dp0data\pairings.js" (
    echo Aucune donnee trouvee. Extraction en cours...
    echo.
    call "%~dp0Mettre a jour les pairings.bat"
)

:: Open the HTML file in default browser
echo Ouverture de l'application...
start "" "%~dp0index.html"
exit
