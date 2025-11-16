# Enable Windows Developer Mode to Fix Build Issues

The build is failing because Windows requires special privileges to create symbolic links (needed for code signing tools). You have two options:

## Option 1: Enable Developer Mode (Recommended - No Admin Needed)

1. Open **Settings** (Windows key + I)
2. Go to **Privacy & Security** → **For developers**
3. Turn on **Developer Mode**
4. Restart your computer (or just restart the terminal)

This allows creating symbolic links without admin privileges.

## Option 2: Run Build as Administrator

1. Right-click on PowerShell or Command Prompt
2. Select **"Run as Administrator"**
3. Navigate to your project folder
4. Run `.\build_electron.bat`

## Option 3: Disable Signing Completely (Current Attempt)

We're trying to disable signing, but electron-builder is still checking for it. After enabling Developer Mode, the build should work even if it tries to download signing tools.

## Quick Fix

**Enable Developer Mode** (takes 2 minutes):
1. Press `Win + I` to open Settings
2. Search for "Developer Mode"
3. Turn it ON
4. Restart terminal/PowerShell
5. Run `.\build_electron.bat` again

This should fix the symbolic link error permanently!

