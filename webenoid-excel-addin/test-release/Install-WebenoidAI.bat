@echo off
echo ============================================
echo    WEBENOID AI - Installing Excel Add-in
echo    (Testing Release Version)
echo ============================================
echo.
echo This will request Administrator permission...
echo.

:: Re-launch as admin if not already
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process cmd -Verb RunAs -ArgumentList '/c \"%~f0\"'"
    exit /b
)

:: Create a folder for the add-in
set "ADDIN_DIR=C:\WebenoidAddin_Test"
if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%"

:: Copy manifest to the folder
copy /Y "%~dp0manifest.xml" "%ADDIN_DIR%\manifest.xml" >nul

:: Create network share (remove old one first if exists)
net share WebenoidAddinTest /delete >nul 2>&1
net share WebenoidAddinTest="%ADDIN_DIR%" /grant:everyone,READ >nul 2>&1

if %errorlevel%==0 (
    echo.
    echo =============================================
    echo   SUCCESS! Webenoid AI has been installed.
    echo =============================================
    echo.
    echo Now do this in Excel:
    echo.
    echo   1. Open Excel
    echo   2. File ^> Options ^> Trust Center ^> Trust Center Settings
    echo   3. Click "Trusted Add-in Catalogs"
    echo   4. Paste:  \\localhost\WebenoidAddinTest
    echo   5. Click "Add catalog" and check "Show in Menu"
    echo   6. Click OK ^> OK
    echo   7. RESTART Excel
    echo   8. Insert ^> My Add-ins ^> SHARED FOLDER tab
    echo   9. Click "Webenoid AI" ^> Add
    echo.
    echo =============================================
) else (
    echo.
    echo ERROR: Could not create network share.
    echo Please try running this file as Administrator.
)

echo.
pause
