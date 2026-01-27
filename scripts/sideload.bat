@echo off
echo ============================================
echo   Outlook AI Assistant - Sideload Helper
echo ============================================
echo.

:: Check if running as admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script needs to run as Administrator.
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: Get the manifest path
set MANIFEST_PATH=%~dp0..\manifest.xml

:: Check if manifest exists
if not exist "%MANIFEST_PATH%" (
    echo ERROR: manifest.xml not found at %MANIFEST_PATH%
    pause
    exit /b 1
)

echo Found manifest at: %MANIFEST_PATH%
echo.

:: Create the sideload directory if it doesn't exist
set SIDELOAD_DIR=%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\Developer
if not exist "%SIDELOAD_DIR%" (
    echo Creating sideload directory...
    mkdir "%SIDELOAD_DIR%"
)

:: Copy manifest to sideload directory  
echo Copying manifest to sideload directory...
copy /Y "%MANIFEST_PATH%" "%SIDELOAD_DIR%\manifest.xml"

:: Enable developer mode in registry
echo Enabling developer mode...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer" /v "EnableDeveloperTools" /t REG_DWORD /d 1 /f

echo.
echo ============================================
echo   Setup Complete!
echo ============================================
echo.
echo Next steps:
echo 1. Make sure dev server is running (npm run dev)
echo 2. Open/Restart Outlook
echo 3. Open an email
echo 4. Look for the AI Assistant in the ribbon or ... menu
echo.
echo If you don't see it, try:
echo - File ^> Options ^> Trust Center ^> Trust Center Settings
echo - Go to "Trusted Add-in Catalogs" 
echo - Check "Enable Developer Add-ins"
echo.
pause
