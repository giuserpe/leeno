@echo off
setlocal

set REPO_PATH=W:\_dwg\ULTIMUSFREE\_SRC\leeno

echo ============================================
echo  LeenO - Ripristina/aggiorna dev
echo  Percorso: %REPO_PATH%
echo ============================================
echo.
echo IMPORTANTE: chiudi LibreOffice prima di continuare,
echo altrimenti alcuni file potrebbero risultare bloccati.
echo.
pause

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
echo  Riapri LibreOffice per riprendere il lavoro.
echo ============================================
pause
