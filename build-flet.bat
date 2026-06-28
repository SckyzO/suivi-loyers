@echo off
REM ===================================================================
REM  POC : construit l'executable Windows de l'interface Flet.
REM  Produit dist\SuiviLoyers.exe (un seul .exe, double-clic).
REM  Build via `flet pack` (PyInstaller + client Flet embarque).
REM  Prerequis : Python 3.10+ (https://www.python.org/downloads/).
REM  N'affecte PAS build.bat (interface Tkinter) : les deux coexistent.
REM ===================================================================
setlocal
cd /d %~dp0

tasklist /FI "IMAGENAME eq SuiviLoyers.exe" 2>NUL | find /I "SuiviLoyers.exe" >NUL
if errorlevel 1 goto :build_start
echo.
echo L'application SuiviLoyers.exe est actuellement ouverte.
set /p REP=La fermer maintenant ? (O/N) :
if /I "%REP%"=="O" (
  REM /T tue l'arbre de process (l'exe onefile lance un bootloader + l'app).
  taskkill /IM SuiviLoyers.exe /F /T >NUL 2>&1
  timeout /t 2 /nobreak >NUL
  echo Application fermee.
) else (
  echo Build annule. Fermez l'application puis relancez build-flet.bat.
  pause
  exit /b 1
)
:build_start

echo [1/3] Creation de l'environnement de build...
python -m venv .buildenv-flet || goto :erreur
call .buildenv-flet\Scripts\activate.bat

echo [2/3] Installation des dependances (Flet + PyInstaller)...
python -m pip install --upgrade pip || goto :erreur
pip install -r requirements-flet.txt pyinstaller || goto :erreur

echo [3/3] Generation de l'executable...
REM Nettoyage prealable : evite l'invite PyInstaller "supprimer build/dist ?".
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
flet pack interface_flet.py ^
  --name SuiviLoyers ^
  --product-name "Suivi des loyers" ^
  --hidden-import yaml ^
  --hidden-import openpyxl || goto :erreur

echo.
echo ===================================================================
echo  Termine. L'executable se trouve ici : dist\SuiviLoyers.exe
echo ===================================================================
echo.
set REP=O
set /p REP=Lancer l'application maintenant ? (O/N) [O] :
if /I not "%REP%"=="N" start "" "dist\SuiviLoyers.exe"
exit /b 0

:erreur
echo.
echo *** Une erreur est survenue pendant la construction. ***
pause
exit /b 1
