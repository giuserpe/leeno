@echo off
setlocal

set REPO_PATH=W:\_dwg\ULTIMUSFREE\_SRC\leeno

echo ============================================
echo  LeenO - Passa a branch Jules per test
echo  Percorso: %REPO_PATH%
echo ============================================
echo.
echo NOTA: LibreOffice puo' restare aperto durante pull/checkout.
echo Va chiuso e riaperto solo se le modifiche toccano la UI
echo (dialoghi .xdl, toolbar/menu, icone) - il codice Python puro
echo viene ricaricato dinamicamente senza bisogno di riavvio.
echo.

cd /d "%REPO_PATH%"
if errorlevel 1 (
    echo ERRORE: impossibile accedere a %REPO_PATH%
    echo Verifica che il drive W: sia collegato.
    pause
    exit /b 1
)

echo Aggiorno riferimenti remoti...
git fetch origin

echo.
set /p BRANCH_NAME="Nome branch Jules da testare (es. jules/nome-task): "

if "%BRANCH_NAME%"=="" (
    echo Nessun branch specificato, esco.
    pause
    exit /b 1
)

git checkout "%BRANCH_NAME%"

if errorlevel 1 (
    echo.
    echo ERRORE: verifica che il nome branch sia corretto
    echo ^(controllalo nella pagina della PR su GitHub^).
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Ora sei su: %BRANCH_NAME%
echo  Se la modifica tocca solo codice Python, LibreOffice
echo  la vede senza riavvio. Per modifiche a UI/dialoghi/
echo  toolbar/icone, chiudi e riapri LibreOffice.
echo.
echo  Al termine del test, lancia leeno_restore_dev.bat
echo  per tornare su dev.
echo ============================================
pause
