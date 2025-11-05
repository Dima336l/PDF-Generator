#!/bin/bash

# Property PDF Builder - macOS Build Script
# This script creates a proper macOS application bundle

echo "🏠 Building Property PDF Builder for macOS..."

# Check if we're on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ This script is designed for macOS only"
    exit 1
fi

# Make sure the script is executable
chmod +x "$0"

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Install requirements
echo "📥 Installing requirements..."
pip install -r requirements.txt

# Clean previous builds
echo "🧹 Cleaning previous builds..."
rm -rf build/
rm -rf dist/
rm -rf PropertyPDFBuilder.app

# Build the macOS app
echo "🔨 Building macOS application..."
pyinstaller PropertyPDFBuilder_macos.spec

# Check if build was successful
if [ -d "dist/PropertyPDFBuilder.app" ]; then
    echo "✅ Build successful!"
    echo "📱 Application created: dist/PropertyPDFBuilder.app"
    echo ""
    echo "🚀 To run the app:"
    echo "   open dist/PropertyPDFBuilder.app"
    echo ""
    echo "📦 To distribute:"
    echo "   - Copy dist/PropertyPDFBuilder.app to Applications folder"
    echo "   - Or create a DMG installer using:"
    echo "     hdiutil create -volname \"Property PDF Builder\" -srcfolder dist/PropertyPDFBuilder.app -ov -format UDZO dist/PropertyPDFBuilder.dmg"
    echo ""
    echo "⚠️  Note: macOS may show a security warning on first run."
    echo "   Right-click the app and select 'Open' to bypass Gatekeeper."
    echo ""
    echo "🔐 Optional: Code sign the app for distribution (requires Apple Developer account):"
    echo "   codesign --deep --force --verify --verbose --sign \"Developer ID Application: Your Name\" dist/PropertyPDFBuilder.app"
    echo "   spctl --assess --verbose dist/PropertyPDFBuilder.app"
else
    echo "❌ Build failed!"
    exit 1
fi

echo "🎉 Done! Your macOS app is ready."
