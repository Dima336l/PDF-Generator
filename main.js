const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'logo.png'),
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    show: false // Don't show until ready
  });
  
  // Show window when ready to prevent white flash
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.loadFile('index.html');

  // Do not auto-open DevTools in production

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();
  
  // Register IPC handlers after window is created

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Register IPC handlers

// Test IPC handler to verify communication works
ipcMain.handle('test-ipc', () => {
  return 'IPC is working';
});

// Handle file dialogs
ipcMain.handle('select-images', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Images', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'] }
    ]
  });
  return result.canceled ? [] : result.filePaths;
});

ipcMain.handle('save-pdf', async (event, defaultFilename) => {
  try {
    // Ensure main window exists and is available
    if (!mainWindow) {
      throw new Error('Main window is not available');
    }
    
    // Ensure main window is focused and visible
    if (mainWindow.isDestroyed()) {
      throw new Error('Main window is destroyed');
    }
    
    mainWindow.focus();
    if (mainWindow.isMinimized()) {
      mainWindow.restore();
    }
    
    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Save Investment Report As',
      defaultPath: defaultFilename,
      filters: [
        { name: 'PDF files', extensions: ['pdf'] }
      ],
      properties: ['showOverwriteConfirmation']
    });
    
    if (result.canceled) {
      return null;
    }
    
    if (!result.filePath) {
      return null;
    }
    
    // Ensure .pdf extension is added if not present
    let filePath = result.filePath;
    if (!filePath.toLowerCase().endsWith('.pdf')) {
      filePath = filePath + '.pdf';
    }
    
    return filePath;
  } catch (error) {
    throw error; // Re-throw so renderer can catch it
  }
});

ipcMain.handle('read-image', async (event, imagePath) => {
  try {
    const imageBuffer = fs.readFileSync(imagePath);
    return imageBuffer;
  } catch (error) {
    return null;
  }
});

ipcMain.handle('get-logo-path', () => {
  // Use path.join for cross-platform compatibility
  const logoPath = path.join(__dirname, 'logo.png');
  // Normalize path separators for consistency
  return fs.existsSync(logoPath) ? path.normalize(logoPath) : null;
});

// Handle getting app directory for sample images
ipcMain.handle('get-app-path', () => {
  return __dirname;
});

