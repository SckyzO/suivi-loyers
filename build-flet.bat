@echo off
REM ===================================================================
REM  POC : construit l'executable Windows de l'interface Flet.
REM  Produit dist\SuiviLoyers.exe (un seul .exe, double-clic).
REM  Build via `flet pack` (PyInstaller + client Flet embarque).
REM  Prerequis : Python 3.10+ (https://www.python.org/downloads/).
REM  N'affecte PAS build.bat (interface Tkinter) : les deux coexistent.
REM
REM  Affichage : un spinner par etape (sortie masquee). Pour voir le
REM  detail complet des commandes, lancer :   build-flet.bat --debug
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

echo.
echo Construction de SuiviLoyers.exe (Flet)   ^(build-flet.bat --debug pour le detail^)
echo.

call :step "Creation de l'environnement (1/3)" venv || goto :erreur
call :step "Installation des dependances (2/3)" deps || goto :erreur
call :step "Generation de l'executable (3/3)"   pack || goto :erreur

echo.
echo ===================================================================
echo  Termine. L'executable se trouve ici : dist\SuiviLoyers.exe
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
set "LOG=%TEMP%\suiviloyers_flet_%NAME%.log"
set "RC=%TEMP%\suiviloyers_flet_%NAME%.rc"
set "RUN=%TEMP%\suiviloyers_flet_%NAME%.cmd"
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
> "%TEMP%\suiviloyers_flet_%NAME%.rc" echo %EC%
exit /b %EC%


REM --- Etapes : les commandes reelles ----------------------------------
:do_venv
python -m venv .buildenv-flet || exit /b 1
exit /b 0

:do_deps
call .buildenv-flet\Scripts\activate.bat || exit /b 1
python -m pip install --upgrade pip || exit /b 1
pip install -r requirements-flet.txt pyinstaller || exit /b 1
exit /b 0

:do_pack
call .buildenv-flet\Scripts\activate.bat || exit /b 1
REM Nettoyage prealable : evite l'invite PyInstaller "supprimer build/dist ?".
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
flet pack interface_flet.py ^
  --name SuiviLoyers ^
  --product-name "Suivi des loyers" ^
  --hidden-import yaml ^
  --hidden-import openpyxl || exit /b 1
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
echo     Relancez avec  build-flet.bat --debug  pour voir le detail.
pause
exit /b 1
