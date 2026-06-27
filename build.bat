@echo off
REM ===================================================================
REM  Construit l'executable Windows autonome : dist\SuiviLoyers.exe
REM  A lancer sous Windows (PyInstaller n'est pas multi-plateforme).
REM  Prerequis : Python 3.10+ installe (https://www.python.org/downloads/).
REM ===================================================================
setlocal
cd /d %~dp0

REM --- Si l'application est ouverte, PyInstaller ne peut pas remplacer le .exe. ---
tasklist /FI "IMAGENAME eq SuiviLoyers.exe" 2>NUL | find /I "SuiviLoyers.exe" >NUL
if errorlevel 1 goto :build_start
echo.
echo L'application SuiviLoyers.exe est actuellement ouverte.
echo Il faut la fermer pour pouvoir reconstruire l'executable.
set /p REP=La fermer maintenant ? (O/N) :
if /I "%REP%"=="O" (
  taskkill /IM SuiviLoyers.exe /F >NUL 2>&1
  echo Application fermee.
) else (
  echo.
  echo Build annule. Fermez l'application puis relancez build.bat.
  pause
  exit /b 1
)
:build_start

echo [1/3] Creation de l'environnement de build...
python -m venv .buildenv || goto :erreur
call .buildenv\Scripts\activate.bat

echo [2/3] Installation des dependances...
python -m pip install --upgrade pip || goto :erreur
pip install -r requirements.txt pyinstaller || goto :erreur

echo [3/3] Generation de l'executable...
pyinstaller --onefile --windowed --noconfirm --clean ^
  --name SuiviLoyers ^
  --collect-submodules openpyxl ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --collect-data sv_ttk ^
  --collect-all darkdetect ^
  --hidden-import babel.numbers ^
  interface.py || goto :erreur

echo.
echo ===================================================================
echo  Termine. L'executable se trouve ici : dist\SuiviLoyers.exe
echo  Vous pouvez le copier ou vous voulez et le lancer par double-clic.
echo ===================================================================
pause
exit /b 0

:erreur
echo.
echo *** Une erreur est survenue pendant la construction. ***
pause
exit /b 1
