@echo off
REM ===================================================================
REM  Construit l'executable Windows autonome : dist\SuiviLoyers.exe
REM  A lancer sous Windows (PyInstaller n'est pas multi-plateforme).
REM  Prerequis : Python 3.10+ installe (https://www.python.org/downloads/).
REM
REM  Affichage : un spinner par etape (sortie masquee). Pour voir le
REM  detail complet des commandes, lancer :   build.bat --debug
REM ===================================================================
setlocal EnableExtensions EnableDelayedExpansion
cd /d %~dp0

REM --- Re-entree interne : execution d'une seule etape en arriere-plan ---
if "%~1"=="__step" goto :__step

REM --- Retour chariot (pour reecrire la ligne du spinner sur place) ------
for /F %%a in ('copy /Z "%~f0" nul') do set "CR=%%a"

REM --- Mode debug : --debug affiche toute la sortie des commandes --------
set "DEBUG="
if /I "%~1"=="--debug" set "DEBUG=1"

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

echo.
echo Construction de SuiviLoyers.exe   ^(build.bat --debug pour le detail^)
echo.

call :step "Creation de l'environnement (1/3)" venv  || goto :erreur
call :step "Installation des dependances (2/3)" deps || goto :erreur
call :step "Generation de l'executable (3/3)"   pack || goto :erreur

echo.
echo ===================================================================
echo  Termine. L'executable se trouve ici : dist\SuiviLoyers.exe
echo  Vous pouvez le copier ou vous voulez et le lancer par double-clic.
echo ===================================================================
echo.
set REP=O
set /p REP=Lancer l'application maintenant ? (O/N) [O] :
if /I not "%REP%"=="N" start "" "dist\SuiviLoyers.exe"
exit /b 0


REM ===================================================================
REM  :step  "Libelle"  nom_etape
REM    debug  -> execute l'etape en clair (sortie visible).
REM    sinon  -> lance l'etape en arriere-plan (sortie -> journal) et
REM              anime un spinner ; en cas d'echec, affiche le journal.
REM ===================================================================
:step
set "LABEL=%~1"
set "NAME=%~2"
if defined DEBUG (
  echo === %LABEL% ===
  call :do_%NAME%
  exit /b %errorlevel%
)
set "LOG=%TEMP%\suiviloyers_%NAME%.log"
set "RC=%TEMP%\suiviloyers_%NAME%.rc"
set "RUN=%TEMP%\suiviloyers_%NAME%.cmd"
del "%LOG%" "%RC%" "%RUN%" 2>NUL
> "%RUN%" echo @call "%~f0" __step %NAME% ^> "%LOG%" 2^>^&1
start "" /b cmd /c "%RUN%"
call :spinner "%LABEL%" "%RC%"
set "CODE=1"
if exist "%RC%" set /p CODE=<"%RC%"
del "%RUN%" 2>NUL
if not "%CODE%"=="0" (
  echo   [ECHEC] %LABEL%
  echo   ----- journal -------------------------------------------------
  if exist "%LOG%" type "%LOG%"
  echo   ---------------------------------------------------------------
  exit /b 1
)
exit /b 0


REM --- Execution effective d'une etape (relancee en arriere-plan) -------
:__step
set "NAME=%~2"
call :do_%NAME%
set "EC=%errorlevel%"
> "%TEMP%\suiviloyers_%NAME%.rc" echo %EC%
exit /b %EC%


REM --- Etapes : les commandes reelles ----------------------------------
:do_venv
python -m venv .buildenv || exit /b 1
exit /b 0

:do_deps
call .buildenv\Scripts\activate.bat || exit /b 1
python -m pip install --upgrade pip || exit /b 1
pip install -r requirements.txt pyinstaller || exit /b 1
exit /b 0

:do_pack
call .buildenv\Scripts\activate.bat || exit /b 1
pyinstaller --onefile --windowed --noconfirm --clean ^
  --name SuiviLoyers ^
  --collect-submodules openpyxl ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --collect-data sv_ttk ^
  --collect-all darkdetect ^
  --hidden-import babel.numbers ^
  interface.py || exit /b 1
exit /b 0


REM --- Spinner : tourne tant que le fichier .rc (fin d'etape) absent ----
:spinner
set "SLABEL=%~1"
set "SRC=%~2"
set "FRAMES=|/-\"
set /a n=0
:spin_loop
if exist "%SRC%" goto :spin_done
set /a fi=n %% 4
for %%# in (!fi!) do set "ch=!FRAMES:~%%#,1!"
<nul set /p "=  !ch!  !SLABEL!... !n!s    !CR!"
set /a n+=1
REM ~1 s sans dependre de stdin (timeout echoue si l'entree est redirigee).
ping -n 2 127.0.0.1 >NUL
goto :spin_loop
:spin_done
<nul set /p "=  [OK] !SLABEL!                                        !CR!"
echo.
exit /b 0


:erreur
echo.
echo *** Une erreur est survenue pendant la construction. ***
echo     Relancez avec  build.bat --debug  pour voir le detail.
pause
exit /b 1
