@echo off
title RFEM HTML zu DOCX Konverter

echo ===================================================
echo RFEM HTML zu DOCX Konverter - Startup
echo ===================================================
echo.

REM Prüfen, ob Python installiert ist
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo FEHLER: Python ist nicht installiert oder nicht im PATH verfügbar.
    echo.
    echo Bitte installieren Sie Python von https://www.python.org/downloads/
    echo und aktivieren Sie die Option "Add Python to PATH" während der Installation.
    echo.
    pause
    exit /b 1
)

echo Python wurde gefunden. Prüfe auf erforderliche Pakete...
echo.

REM Prüfen und installieren der erforderlichen Pakete
pip install beautifulsoup4 requests pillow cairosvg python-docx --quiet

echo.
echo Starte RFEM HTML zu DOCX Konverter...
echo.

REM Programm starten
python rfem6_html_converter.py

REM Falls ein Fehler auftritt, zeige Meldung an
if %errorlevel% neq 0 (
    echo.
    echo Es ist ein Fehler beim Starten des Programms aufgetreten.
    echo Bitte stellen Sie sicher, dass alle Abhängigkeiten korrekt installiert sind.
    echo.
    echo Fehlercode: %errorlevel%
    pause
)

exit /b 0
