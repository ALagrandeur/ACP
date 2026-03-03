@echo off
chcp 65001 >nul
title Mise a jour des Pairings

echo ============================================
echo    Mise a jour des Pairings
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Python n'est pas installe ou pas dans le PATH.
    echo Installez Python depuis https://python.org
    pause
    exit /b 1
)

:: Check pymupdf
python -c "import fitz" >nul 2>&1
if errorlevel 1 (
    echo Installation de PyMuPDF...
    python -m pip install pymupdf
    echo.
)

:: Run extraction
echo Recherche du PDF le plus recent dans le dossier Pairing...
echo.
python "%~dp0extract_pairings.py"

if errorlevel 1 (
    echo.
    echo ERREUR: L'extraction a echoue.
    echo Verifiez que votre PDF est dans le dossier:
    echo   %~dp0Pairing\
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================
echo    Mise a jour terminee!
echo ============================================
echo.
echo Pour ouvrir l'application, double-cliquez sur:
echo   "Ouvrir Pairings.bat"
echo.
pause
