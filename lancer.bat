@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=python
%PY% --version >nul 2>&1
if errorlevel 1 set PY=py
%PY% --version >nul 2>&1
if errorlevel 1 goto sanspython
%PY% pronote_suivi.py %*
goto fin
:sanspython
echo Python n'est pas installe sur cet ordinateur.
echo Ouvrez le fichier LISEZMOI.html et suivez l'etape "Installer Python".
:fin
echo.
pause
