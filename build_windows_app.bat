@echo off
REM Property PDF Builder - Windows Build Script
REM This script creates a Windows executable

echo 🏠 Building Property PDF Builder for Windows...

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo 📦 Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo 🔧 Activating virtual environment...
call venv\Scripts\activate.bat

REM Install requirements
echo 📥 Installing requirements...
pip install -r requirements.txt

REM Clean previous builds
echo 🧹 Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "PropertyPDFBuilder.exe" del PropertyPDFBuilder.exe

REM Build the Windows executable
echo 🔨 Building Windows executable...
pyinstaller --onefile --windowed --name PropertyPDFBuilder pdf_builder_app.py

REM Check if build was successful
if exist "dist\PropertyPDFBuilder.exe" (
    echo ✅ Build successful!
    echo 📱 Executable created: dist\PropertyPDFBuilder.exe
    echo.
    echo 🚀 To run the app:
    echo    dist\PropertyPDFBuilder.exe
) else (
    echo ❌ Build failed!
    exit /b 1
)

echo 🎉 Done! Your Windows app is ready.
pause
