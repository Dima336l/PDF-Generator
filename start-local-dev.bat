@echo off
REM Start script for local development
REM This starts both the backend server and the web frontend

echo ========================================
echo Starting Local Development Environment
echo ========================================
echo.

REM Function to kill process on a specific port
echo Checking for existing processes on ports 3000 and 8080...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":3000" ^| findstr "LISTENING"') do @taskkill /F /PID %%a >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8080" ^| findstr "LISTENING"') do @taskkill /F /PID %%a >nul 2>&1

REM Wait a moment for ports to be released
timeout /t 1 /nobreak >nul
echo.

REM Check if backend node_modules exist
if not exist "backend\node_modules" (
    echo Backend dependencies not found. Installing...
    cd backend
    call npm install
    cd ..
    echo.
)

REM Check if root node_modules exist
if not exist "node_modules" (
    echo Frontend dependencies not found. Installing...
    call npm install
    echo.
)

echo Starting backend server on port 8080...
start "Backend Server" cmd /k "cd backend && npm start"

REM Wait a moment for backend to start
timeout /t 2 /nobreak >nul

echo Starting web server...
echo.
echo Backend is running at: http://localhost:8080
echo Frontend will be available at: http://localhost:3000
echo.
echo Open your browser and navigate to: http://localhost:3000
echo.
echo To stop: Press Ctrl+C in this window
echo.

call npm start

