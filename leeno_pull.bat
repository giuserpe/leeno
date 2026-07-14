@echo off
setlocal

set REPO_PATH=W:\_dwg\ULTIMUSFREE\_SRC\leeno

echo ============================================
echo  LeenO - Pull branch dev
echo  Percorso: %REPO_PATH%
echo ============================================
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
echo  Fatto.
echo ============================================
pause
