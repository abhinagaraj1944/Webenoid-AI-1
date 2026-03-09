@echo off
echo ============================================
echo    WEBENOID AI - Uninstalling Add-in
echo ============================================
echo.

set "ADDIN_DIR=C:\WebenoidAddin_Test"
net share WebenoidAddinTest /delete >nul 2>&1

if exist "%ADDIN_DIR%" (
    rmdir /S /Q "%ADDIN_DIR%"
    echo.
    echo Webenoid AI testing release has been removed.
    echo Please restart Excel.
) else (
    echo Webenoid AI testing release was not found.
)

echo.
pause
