# macOS Installation Instructions

## First-Time Setup

Since this app is not code-signed, macOS Gatekeeper will prevent it from opening on first launch. Follow these steps:

### Option 1: Right-Click to Open (Recommended)
1. Download the `PropertyPDFBuilder.dmg` file
2. Double-click the DMG to mount it
3. Drag `PropertyPDFBuilder.app` to your Applications folder
4. **Right-click** on `PropertyPDFBuilder.app` in Applications
5. Select **"Open"** from the context menu
6. Click **"Open"** in the security dialog that appears
7. The app will now open normally, and you can double-click it in the future

### Option 2: Remove Quarantine Attribute (Terminal)
If Option 1 doesn't work, open Terminal and run:

```bash
xattr -dr com.apple.quarantine /Applications/PropertyPDFBuilder.app
```

Then you can double-click the app normally.

### Option 3: System Settings
1. Go to **System Settings** → **Privacy & Security**
2. Scroll down to find a message about "PropertyPDFBuilder" being blocked
3. Click **"Open Anyway"**

## Troubleshooting

**"The application can't be opened"**
- Use Option 1 (right-click → Open) instead of double-clicking
- Or use Option 2 to remove quarantine attributes

**App won't launch after following instructions**
- Check that you're running macOS 10.13 or later
- Try moving the app to a different location (Desktop, Documents)
- Check Console.app for error messages

## System Requirements
- macOS 10.13 (High Sierra) or later
- 50 MB free disk space

