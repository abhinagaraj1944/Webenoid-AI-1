@echo off
echo ============================================
echo   WEBENOID AI - Uninstalling Add-in
echo ============================================
echo.

:: Auto Re-launch as Admin if not already
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process cmd -Verb RunAs -ArgumentList '/c \"%~f0\"'"
    exit /b
)

set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

if exist "%WEF%\webenoid-manifest.xml" (
    del /f "%WEF%\webenoid-manifest.xml"
    echo Webenoid AI has been removed.
    echo Please restart Excel.
) else (
    echo Webenoid AI was not found. Nothing to remove.
)

echo.
pause
