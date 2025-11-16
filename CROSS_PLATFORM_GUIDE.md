# Cross-Platform Development Guide

This guide explains how the Electron app ensures compatibility between Windows and macOS.

## Single Codebase Principle

**The entire application uses ONE codebase that works on both platforms.** There are no platform-specific versions needed.

## How It Works

### 1. **Node.js Path Module**
All file paths use Node.js's `path` module which automatically handles:
- Windows: `C:\Users\...` (backslashes)
- macOS: `/Users/...` (forward slashes)

```javascript
const path = require('path');
const logoPath = path.join(__dirname, 'logo.png'); // Works on both platforms
```

### 2. **Electron APIs**
Electron's APIs are cross-platform by default:
- File dialogs work the same on both platforms
- Window management is consistent
- IPC (Inter-Process Communication) works identically

### 3. **File URL Handling**
The code normalizes file paths for `file://` URLs to work on both platforms:

```javascript
// Windows: C:\path\to\file.jpg → file:///C:/path/to/file.jpg
// macOS: /path/to/file.jpg → file:///path/to/file.jpg
```

### 4. **Build Process**
You can build for both platforms from either platform:

**From Windows:**
```bash
npm run build:win   # Builds Windows .exe
npm run build:mac   # Requires macOS (or use CI/CD)
npm run build:all   # Builds both (macOS build requires Mac)
```

**From macOS:**
```bash
npm run build:win   # Builds Windows .exe
npm run build:mac   # Builds macOS .app
npm run build:all   # Builds both
```

## Testing Strategy

### Recommended Approach

1. **Develop on Windows** (or macOS - doesn't matter)
2. **Test locally** on your development platform
3. **Build for both platforms**:
   - If on Windows: Build Windows version, test it
   - If on macOS: Build macOS version, test it
   - Use CI/CD (GitHub Actions) to build the other platform

### GitHub Actions (Recommended)

Set up automated builds so you don't need both platforms:

```yaml
# .github/workflows/build.yml
- name: Build Windows
  if: runner.os == 'Windows'
  run: npm run build:win

- name: Build macOS  
  if: runner.os == 'macOS'
  run: npm run build:mac
```

This way, every push automatically builds both platforms.

## What's Already Cross-Platform

✅ **File paths** - Using `path.join()` and `path.normalize()`  
✅ **File dialogs** - Electron handles platform differences  
✅ **File operations** - `fs` module works identically  
✅ **Image loading** - Normalized file:// URLs  
✅ **PDF generation** - PDFKit works on both platforms  
✅ **UI rendering** - HTML/CSS/JS is platform-agnostic  

## Platform-Specific Considerations

### Only One Platform-Specific Feature

The only platform-specific code is the window title bar style:

```javascript
titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default'
```

This is cosmetic and doesn't affect functionality.

### Building Requirements

- **Windows builds** can be done on Windows or macOS
- **macOS builds** typically require macOS (or use CI/CD)
- **Linux builds** can be added if needed

## Verification Checklist

Before releasing, verify:

- [ ] App runs on Windows
- [ ] App runs on macOS  
- [ ] File dialogs work on both
- [ ] Image loading works on both
- [ ] PDF generation works on both
- [ ] File paths are handled correctly
- [ ] No hardcoded path separators (`\` or `/`)

## Common Issues & Solutions

### Issue: Images not loading on Windows
**Solution:** Already fixed - file:// URLs are normalized in `renderer.js`

### Issue: Path errors
**Solution:** Always use `path.join()` instead of string concatenation

### Issue: Build fails on different platform
**Solution:** Use GitHub Actions or build on the target platform

## Best Practices

1. **Always use `path.join()`** for file paths
2. **Never hardcode** `\` or `/` separators
3. **Test file dialogs** on both platforms
4. **Use Electron's built-in APIs** (they're cross-platform)
5. **Build and test** on both platforms before release

## Summary

**You have ONE codebase that works on BOTH platforms.** The code is already set up to handle platform differences automatically. Just build for the platform you want, and it will work!

