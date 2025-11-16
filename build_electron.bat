@echo off

REM Build script for Electron app (Windows)
REM NOTE: If you get symbolic link errors, enable Windows Developer Mode:
REM Settings > Privacy & Security > For developers > Turn ON Developer Mode

echo Installing dependencies...
call npm install

echo Cleaning previous build...
if exist dist-electron rmdir /s /q dist-electron

echo Cleaning electron-builder cache (to remove corrupted signing tools)...
if exist "%LOCALAPPDATA%\electron-builder\Cache\winCodeSign" rmdir /s /q "%LOCALAPPDATA%\electron-builder\Cache\winCodeSign"
if exist "%LOCALAPPDATA%\electron-builder\Cache" rmdir /s /q "%LOCALAPPDATA%\electron-builder\Cache"

echo Building Electron app for Windows...
REM Disable code signing to avoid privilege issues
set CSC_IDENTITY_AUTO_DISCOVERY=false
call npm run build:win

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo BUILD FAILED - Symbolic Link Error?
    echo ========================================
    echo.
    echo Enable Windows Developer Mode to fix this:
    echo 1. Press Win+I to open Settings
    echo 2. Go to Privacy ^& Security ^> For developers
    echo 3. Turn ON Developer Mode
    echo 4. Restart terminal and try again
    echo.
    echo Or run this script as Administrator
    echo.
) else (
    echo.
    echo Build complete! Check the dist-electron folder for your executables.
    echo The app is in: dist-electron\win-unpacked\
    echo Run electron.exe to launch the app.
)

pause

