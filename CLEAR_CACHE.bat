@echo off
REM Clear electron-builder cache to fix code signing issues

echo Clearing electron-builder cache...
if exist "%LOCALAPPDATA%\electron-builder\Cache" (
    echo Found cache directory, removing...
    rmdir /s /q "%LOCALAPPDATA%\electron-builder\Cache"
    echo Cache cleared!
) else (
    echo Cache directory not found.
)

pause

