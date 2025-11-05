# GitHub Actions Setup Guide

This guide will walk you through setting up GitHub Actions to automatically build both Windows and macOS executables.

## Prerequisites

1. **GitHub Account** - You need a GitHub account (free is fine)
2. **Repository** - Your code should be in a GitHub repository
3. **Git** - Make sure you have git installed locally

## Step-by-Step Setup

### Step 1: Push Your Code to GitHub

If you haven't already pushed your code:

```bash
# Check if you have a remote repository
git remote -v

# If no remote exists, add your GitHub repository
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git

# Push your code
git add .
git commit -m "Add GitHub Actions workflow for cross-platform builds"
git push -u origin main
```

### Step 2: Verify Workflow Files Are Committed

Make sure the workflow file is in your repository:

```bash
# Check if the workflow file exists
ls .github/workflows/

# If it doesn't exist locally, commit it
git add .github/workflows/build-all.yml
git commit -m "Add cross-platform build workflow"
git push
```

### Step 3: Enable GitHub Actions

1. **Go to your GitHub repository** in a web browser
2. **Click on the "Settings" tab** (top menu)
3. **Click "Actions"** in the left sidebar
4. **Under "Actions permissions"**, make sure:
   - ✅ "Allow all actions and reusable workflows" is selected
   - Or "Allow local actions and reusable workflows" if you prefer

### Step 4: Trigger the Workflow

#### Option A: Automatic Trigger (Recommended)
- The workflow will **automatically run** when you push to the `main` branch
- Just push your code and it will build automatically

#### Option B: Manual Trigger
1. Go to the **"Actions"** tab in your GitHub repository
2. Click **"Build All Platforms"** in the left sidebar
3. Click **"Run workflow"** button (top right)
4. Select the branch (usually `main`)
5. Click **"Run workflow"**

### Step 5: Monitor the Build

1. **Go to the "Actions" tab**
2. **Click on the workflow run** (it will show "Build All Platforms" with a yellow/orange dot while running)
3. **Watch the progress**:
   - You'll see two jobs: "Build Windows App" and "Build macOS App"
   - Each job shows steps as they execute
   - Green checkmark = success
   - Red X = failure

### Step 6: Download the Built Apps

Once the workflow completes (usually 5-10 minutes):

1. **Click on the completed workflow run**
2. **Scroll down to the "Artifacts" section**
3. **Download the artifacts**:
   - **PropertyPDFBuilder-Windows** - Contains `PropertyPDFBuilder.exe`
   - **PropertyPDFBuilder-macOS** - Contains `PropertyPDFBuilder.app` and `PropertyPDFBuilder.dmg`

4. **Extract the downloaded zip files** to get your executables

## Troubleshooting

### Workflow Not Running?

1. **Check Actions is enabled**:
   - Go to Settings → Actions
   - Make sure Actions are enabled

2. **Check workflow file syntax**:
   - Go to Actions tab
   - If there's a red X, click it to see the error
   - Common issues: YAML syntax errors, missing files

3. **Check file paths**:
   - Make sure `PropertyPDFBuilder_macos.spec` exists
   - Make sure `requirements.txt` exists
   - Make sure `pdf_builder_app.py` exists

### Build Fails?

1. **Check the logs**:
   - Click on the failed job
   - Expand the failed step to see error messages
   - Common issues:
     - Missing dependencies
     - File path issues
     - Python version incompatibility

2. **Common fixes**:
   - Update `requirements.txt` if dependencies are missing
   - Check that all required files are committed
   - Verify Python version in workflow matches your code

### Can't Download Artifacts?

- Artifacts are available for **30 days** after the build
- Make sure you're logged into GitHub
- Try refreshing the page
- Check if the artifact step completed successfully

## Workflow Features

The workflow I created:
- ✅ Builds **Windows** executable (`.exe`)
- ✅ Builds **macOS** app bundle (`.app`)
- ✅ Creates **macOS DMG** installer
- ✅ Runs on every push to `main`
- ✅ Can be triggered manually
- ✅ Stores artifacts for 30 days

## Customization

### Change When It Runs

Edit `.github/workflows/build-all.yml`:

```yaml
on:
  push:
    branches: [ main, develop ]  # Add more branches
  workflow_dispatch:  # Manual trigger
  schedule:  # Run on schedule
    - cron: '0 0 * * 0'  # Every Sunday at midnight
```

### Change Python Version

Edit `.github/workflows/build-all.yml`:

```yaml
python-version: '3.11'  # Change to '3.10', '3.12', etc.
```

### Add Code Signing (macOS)

Add this step after building:

```yaml
- name: Code sign macOS app
  run: |
    codesign --deep --force --verify --verbose --sign "Developer ID Application: Your Name" dist/PropertyPDFBuilder.app
  env:
    APPLE_CERTIFICATE: ${{ secrets.APPLE_CERTIFICATE }}
    APPLE_CERTIFICATE_PASSWORD: ${{ secrets.APPLE_CERTIFICATE_PASSWORD }}
```

## Next Steps

1. **Test the workflow** - Push a small change and watch it build
2. **Download the artifacts** - Verify the executables work
3. **Share the executables** - Distribute the `.exe` and `.app` files to users
4. **Set up releases** - Optionally create GitHub Releases for easier distribution

## Need Help?

- Check GitHub Actions documentation: https://docs.github.com/en/actions
- Check workflow logs for specific error messages
- Verify all files are committed and pushed to GitHub

