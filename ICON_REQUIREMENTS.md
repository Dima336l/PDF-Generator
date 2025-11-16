# Icon Requirements

## Current Issue

Electron Builder requires app icons to be at least **256x256 pixels**. Your current `logo.png` is smaller than this requirement.

## Solutions

### Option 1: Resize Your Logo (Recommended)

1. Open `logo.png` in an image editor
2. Resize it to at least 256x256 pixels (square format recommended)
3. Save it as `logo.png` (replace the existing file)
4. Rebuild the app

**Recommended sizes:**
- **Windows**: 256x256 (minimum), 512x512 (recommended)
- **macOS**: 512x512 (minimum), 1024x1024 (recommended)

### Option 2: Create Platform-Specific Icons

You can create separate icon files for each platform:

**For Windows:**
- Create `build/icon.ico` (256x256 or larger, in ICO format)
- Or create `build/icon.png` (256x256 or larger)

**For macOS:**
- Create `build/icon.icns` (512x512 or larger, in ICNS format)
- Or create `build/icon.png` (512x512 or larger)

Then update `package.json`:

```json
"win": {
  "icon": "build/icon.ico"
},
"mac": {
  "icon": "build/icon.icns"
}
```

### Option 3: Use Online Icon Generator

1. Go to https://www.electron.build/icons or similar tool
2. Upload your logo
3. Download the generated icons
4. Place them in the appropriate locations

### Option 4: Build Without Custom Icon (Current Setup)

The build configuration has been updated to not require an icon. The app will use Electron's default icon. This works fine for development and testing, but you'll want a custom icon for distribution.

## Quick Fix

If you just want to build now without fixing the icon:

1. The build config has been updated to skip the icon requirement
2. Run `npm run build:win` again
3. The app will build successfully with a default icon

## For Production

Before distributing your app, you should:
1. Create a proper 256x256+ icon
2. Update the build config to use it
3. Rebuild the app

This ensures your app has a professional appearance in the Start Menu, Dock, and file explorer.

