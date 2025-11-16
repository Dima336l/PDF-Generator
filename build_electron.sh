#!/bin/bash

# Build script for Electron app (macOS and Windows)

echo "Installing dependencies..."
npm install

echo "Cleaning previous build..."
rm -rf dist-electron

echo "Building Electron app for all platforms..."
# Disable code signing to avoid issues
export CSC_IDENTITY_AUTO_DISCOVERY=false
npm run build:all

echo "Build complete! Check the dist-electron folder for your executables."

