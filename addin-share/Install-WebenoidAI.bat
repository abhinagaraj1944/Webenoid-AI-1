@echo off
echo ============================================
echo   WEBENOID AI - Installing Excel Add-in
echo ============================================
echo.

:: Auto Re-launch as Admin if not already
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator permission...
    powershell -Command "Start-Process cmd -Verb RunAs -ArgumentList '/c \"%~f0\"'"
    exit /b
)

:: Define the correct Wef path
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

:: Create Wef folder if it doesn't exist
if not exist "%WEF%" mkdir "%WEF%"

:: Copy the manifest directly to the Wef folder
copy /Y "%~dp0manifest.xml" "%WEF%\webenoid-manifest.xml" >nul 2>&1

if %errorlevel%==0 (
    echo.
    echo =============================================
    echo   SUCCESS! Webenoid AI has been installed.
    echo =============================================
    echo.
    echo NEXT STEPS:
    echo.
    echo   1. Close ALL Excel windows completely
    echo   2. Reopen Excel
    echo   3. Click Insert ^> My Add-ins
    echo   4. Click "Webenoid AI" then press Add
    echo.
    echo =============================================
    echo   Enjoy your AI Excel Copilot!
    echo =============================================
) else (
    echo.
    echo ERROR: Installation failed.
    echo Please right-click the file and choose
    echo "Run as administrator".
)

echo.
pause
