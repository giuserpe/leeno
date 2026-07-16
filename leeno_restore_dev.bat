@echo off
setlocal

set REPO_PATH=W:\_dwg\ULTIMUSFREE\_SRC\leeno

echo ============================================
echo  LeenO - Ripristina/aggiorna dev
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

git checkout dev
git pull

echo.
echo ============================================
echo  Ora sei su: dev, aggiornato all'ultimo commit.
echo  Se necessario (modifiche a UI/dialoghi/toolbar/icone),
echo  chiudi e riapri LibreOffice per vedere gli aggiornamenti.
echo ============================================
pause
