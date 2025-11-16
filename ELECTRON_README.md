# Property PDF Builder - Electron Version

This is the Electron-based version of the Property PDF Builder, which allows you to build the same application for both Windows (.exe) and macOS (.app) from a **single codebase**.

## 🎯 Single Codebase, Multiple Platforms

**One codebase works on both Windows and macOS.** There are no separate versions to maintain. The code automatically handles platform differences (file paths, UI styling, etc.) so you can develop on one platform and deploy to both.

## Features

- **Cross-Platform**: Build for both Windows and macOS from the same code
- **No Python Required**: Uses Node.js and Electron instead
- **Native Feel**: Electron provides a native desktop experience
- **Easy Distribution**: Single executable files for each platform

## Prerequisites

- **Node.js** (v16 or higher) - [Download here](https://nodejs.org/)
- **npm** (comes with Node.js)

## Installation

1. **Install dependencies**:
   ```bash
   npm install
   ```

## Development

To run the app in development mode:

```bash
npm start
```

## Building

### Build for All Platforms

```bash
npm run build:all
```

This will create:
- Windows: `.exe` installer and portable executable in `dist-electron/`
- macOS: `.dmg` installer and `.zip` archive in `dist-electron/`

### Build for Specific Platform

**Windows only:**
```bash
npm run build:win
```

**macOS only:**
```bash
npm run build:mac
```

### Using Build Scripts

**On macOS/Linux:**
```bash
chmod +x build_electron.sh
./build_electron.sh
```

**On Windows:**
```cmd
build_electron.bat
```

## Output

After building, you'll find the executables in the `dist-electron/` folder:

- **Windows**: 
  - `Property PDF Builder Setup X.X.X.exe` (installer)
  - `Property PDF Builder X.X.X.exe` (portable)
  
- **macOS**:
  - `Property PDF Builder-X.X.X.dmg` (installer)
  - `Property PDF Builder-X.X.X-mac.zip` (archive)

## Project Structure

```
Property-PDF/
├── main.js              # Electron main process
├── index.html           # Application UI
├── styles.css           # Application styles
├── renderer.js          # UI logic and event handling
├── pdf-generator.js     # PDF generation logic
├── package.json         # Node.js dependencies and build config
├── logo.png             # Application logo
└── dist-electron/       # Build output (created after build)
```

## How It Works

1. **Main Process** (`main.js`): Handles window creation, file dialogs, and system integration
2. **Renderer Process** (`renderer.js`): Manages the UI, form handling, and image management
3. **PDF Generator** (`pdf-generator.js`): Converts form data and images into PDF documents using PDFKit

## Differences from Python Version

- Uses **PDFKit** instead of ReportLab for PDF generation
- Uses **Sharp** for image processing instead of PIL
- HTML/CSS/JavaScript UI instead of tkinter
- **Single codebase for both platforms** - no platform-specific code needed
- **Cross-platform by default** - file paths, dialogs, and all features work identically on Windows and macOS

## Cross-Platform Guarantee

The codebase is designed to work identically on both platforms:

✅ **File paths** - Automatically handles Windows (`C:\`) and macOS (`/`) paths  
✅ **File dialogs** - Native dialogs on each platform  
✅ **Image loading** - Works with both path formats  
✅ **PDF generation** - Identical output on both platforms  
✅ **UI/UX** - Consistent experience across platforms  

**See `CROSS_PLATFORM_GUIDE.md` for detailed information.**

## Troubleshooting

### Build Fails

- Make sure Node.js v16+ is installed: `node --version`
- Delete `node_modules` and `package-lock.json`, then run `npm install` again
- On macOS, you may need to allow the build process in System Preferences > Security

### App Won't Start

- Check that all dependencies are installed: `npm install`
- Try running in development mode first: `npm start`
- Check the console for error messages

### Images Not Loading

- Ensure image paths are valid
- Supported formats: JPG, PNG, GIF, BMP, WEBP
- Check file permissions

## Distribution

### Windows

The `.exe` file is standalone and can be distributed directly. Users can:
- Run the installer to install the app
- Or use the portable `.exe` version

### macOS

The `.dmg` file is the standard macOS installer. Users can:
- Double-click to mount the DMG
- Drag the app to Applications folder
- On first launch, users may need to right-click and select "Open" to bypass Gatekeeper

## License

Same as the original project (MIT License).

