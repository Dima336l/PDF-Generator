# Local Development Setup

This guide will help you run the Property PDF Builder application locally for testing and development.

## Prerequisites

- Node.js (v14 or higher)
- npm (comes with Node.js)

## Quick Start

### Option 1: Using the Batch Script (Windows - Recommended)

Simply run:
```bash
start-local-dev.bat
```

This script will:
1. Check and install dependencies if needed
2. Start the backend server on port 8080
3. Start the Electron frontend

### Option 2: Manual Start

#### Step 1: Install Backend Dependencies

```bash
cd backend
npm install
```

#### Step 2: Start the Backend Server

```bash
cd backend
npm start
```

The backend will start on `http://localhost:8080`

#### Step 3: Install Frontend Dependencies (if not already done)

```bash
npm install
```

#### Step 4: Start the Frontend

In a new terminal window:
```bash
npm start
```

This will launch the Electron application.

### Option 3: Using npm Scripts (with concurrently)

If you prefer to use npm scripts:

```bash
# Install dependencies first
cd backend && npm install && cd ..
npm install

# Start both backend and frontend together
npm run start:dev
```

## Backend Configuration

The backend server runs on port **8080** by default. You can change this by setting the `PORT` environment variable:

```bash
# Windows
set PORT=3000
npm start

# Linux/Mac
PORT=3000 npm start
```

## Frontend Configuration

The frontend automatically detects if it's running locally and will use `http://localhost:8080` for the backend API. 

If you need to manually override the backend URL, you can set it in the browser console:

```javascript
localStorage.setItem('backend_url', 'http://localhost:8080');
```

## Testing the Setup

1. **Backend Health Check**: Open `http://localhost:8080/health` in your browser. You should see `{"ok":true}`

2. **Frontend**: The Electron app should open automatically. Fill in the form and try generating a PDF.

## Troubleshooting

### Backend won't start
- Check if port 8080 is already in use: `netstat -ano | findstr :8080` (Windows)
- Make sure all dependencies are installed: `cd backend && npm install`

### Frontend can't connect to backend
- Verify backend is running: Check `http://localhost:8080/health`
- Check the browser console for CORS errors
- Make sure the backend URL in `renderer.js` is correct

### Dependencies issues
- Delete `node_modules` folders and `package-lock.json` files
- Run `npm install` again in both root and backend directories

## Development vs Production

- **Local Development**: Frontend uses `http://localhost:8080`
- **Production/Hosted**: Frontend uses `https://pdf-generator-backend-fbtb.onrender.com`

The frontend automatically detects the environment and uses the appropriate backend URL.

## Stopping the Services

- **Backend**: Press `Ctrl+C` in the backend terminal window
- **Frontend**: Close the Electron window
- **Batch Script**: Close both terminal windows

