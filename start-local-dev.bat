@echo off
REM Start script for local development
REM This starts both the backend server and the Electron frontend

echo ========================================
echo Starting Local Development Environment
echo ========================================
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

echo Starting Electron frontend...
echo.
echo Backend is running at: http://localhost:8080
echo Frontend will open in Electron window
echo.
echo To stop: Close both windows or press Ctrl+C in each
echo.

call npm start

