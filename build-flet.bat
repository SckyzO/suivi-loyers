@echo off
REM ===================================================================
REM  POC : construit l'executable Windows de l'interface Flet.
REM  Produit dist\SuiviLoyers.exe (un seul .exe, double-clic).
REM  Build via `flet pack` (PyInstaller + client Flet embarque).
REM  Prerequis : Python 3.10+ (https://www.python.org/downloads/)
REM             et uv (https://docs.astral.sh/uv/ ; `pip install uv`).
REM  N'affecte PAS build.bat (interface Tkinter) : les deux coexistent.
REM
REM  Affichage : une ligne "[OK]" par etape (sortie masquee -> journal, montre
REM  seulement en cas d'echec). Detail complet :   build-flet.bat --debug
REM ===================================================================
setlocal EnableExtensions EnableDelayedExpansion
cd /d %~dp0

REM --- Mode debug : --debug affiche toute la sortie des commandes --------
set "DEBUG="
if /I "%~1"=="--debug" set "DEBUG=1"

tasklist /FI "IMAGENAME eq SuiviLoyers.exe" 2>NUL | find /I "SuiviLoyers.exe" >NUL
if errorlevel 1 goto :build_start
echo.
echo L'application SuiviLoyers.exe est actuellement ouverte.
REM Defaut = Oui (Entree ferme) ; accepte O/Oui/N/Non (on teste la 1re lettre).
set "REP=O"
set /p REP=La fermer maintenant ? (Oui/Non) [Oui] :
if /I "%REP:~0,1%"=="O" (
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
REM Defaut = Oui ; accepte O/Oui/N/Non (on teste la 1re lettre).
set "REP=O"
set /p REP=Lancer l'application maintenant ? (Oui/Non) [Oui] :
if /I not "%REP:~0,1%"=="N" start "" "dist\SuiviLoyers.exe"
exit /b 0


REM ===================================================================
REM  :step  "Libelle"  nom_etape
REM    debug  -> execute l'etape en clair (sortie visible).
REM    sinon  -> execute l'etape en redirigeant TOUTE sa sortie (y compris
REM              les sous-processus de flet pack) vers un journal, et n'affiche
REM              ce journal qu'en cas d'echec. Sortie propre : une ligne [OK].
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
del "%LOG%" 2>NUL
<nul set /p "=  %LABEL%... "
REM Rediriger le `call` capture aussi la sortie des programmes qu'il lance.
call :do_%NAME% > "%LOG%" 2>&1
if errorlevel 1 (
  echo [ECHEC]
  echo   ----- journal -------------------------------------------------
  if exist "%LOG%" type "%LOG%"
  echo   ---------------------------------------------------------------
  exit /b 1
)
echo [OK]
exit /b 0


REM --- Etapes : les commandes reelles ----------------------------------
:do_venv
REM --clear : repart d'un environnement vierge a chaque build. Sans ca, un
REM flet_desktop d'une version precedente subsiste et casse `flet pack`
REM (cannot import name 'ensure_client_cached' from 'flet_desktop').
uv venv -q --clear .buildenv-flet || exit /b 1
exit /b 0

:do_deps
call .buildenv-flet\Scripts\activate.bat || exit /b 1
uv pip install -q -r requirements-flet.txt pyinstaller || exit /b 1
exit /b 0

:do_pack
call .buildenv-flet\Scripts\activate.bat || exit /b 1
REM Nettoyage prealable : evite l'invite PyInstaller "supprimer build/dist ?".
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
REM --log-level=WARN : coupe les ~90 lignes INFO de PyInstaller (sortie propre).
flet pack interface_flet.py ^
  --name SuiviLoyers ^
  --product-name "Suivi des loyers" ^
  --hidden-import yaml ^
  --hidden-import openpyxl ^
  --pyinstaller-build-args="--log-level=WARN" || exit /b 1
exit /b 0


:erreur
echo.
echo *** Une erreur est survenue pendant la construction. ***
echo     Relancez avec  build-flet.bat --debug  pour voir le detail.
pause
exit /b 1
