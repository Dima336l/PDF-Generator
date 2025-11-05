# Property PDF Builder - Desktop Application

A professional desktop application that automatically generates property PDFs with images and property details. Available for both **macOS** and **Windows**.

## Features

- **Native Desktop App**: Proper `.app` bundle for Mac or `.exe` for Windows
- **Easy-to-use GUI**: Clean interface built with tkinter
- **Property Information Forms**: Input fields for address, property type, bedrooms, bathrooms, square footage, price, and description
- **Image Management**: Add multiple property images with drag-and-drop ordering
- **Live Preview**: See how your PDF will look before generating
- **Professional PDF Output**: Clean, formatted PDFs with tables, images, and proper styling
- **Cross-Platform**: Works on both macOS and Windows

## Quick Start

### For macOS Users:

1. **Clone or download** this repository
2. **Run the build script**:
   ```bash
   ./build_macos_app.sh
   ```
3. **Launch the app**:
   ```bash
   open dist/PropertyPDFBuilder.app
   ```

### For Windows Users:

1. **Clone or download** this repository
2. **Run the build script**:
   ```cmd
   build_windows_app.bat
   ```
3. **Launch the app**:
   ```cmd
   dist\PropertyPDFBuilder.exe
   ```

## Manual Installation (Alternative)

If you prefer to run from source:

1. **Install Python** (3.7 or higher) from [python.org](https://python.org)

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate virtual environment**:
   - **macOS/Linux**: `source venv/bin/activate`
   - **Windows**: `venv\Scripts\activate`

4. **Install required packages**:
   ```bash
   pip install -r requirements.txt
   ```

5. **Run the application**:
   ```bash
   python pdf_builder_app.py
   ```

## How to Use

1. **Fill in Property Information**:
   - Enter the property address
   - Add postal code, property type, bedrooms, bathrooms, square footage, and price
   - Write a description of the property

2. **Add Images**:
   - Click "Add Images" to select property photos
   - Use "Move Up" and "Move Down" to reorder images
   - Remove unwanted images with "Remove Selected"

3. **Preview**:
   - The preview section shows how your PDF will look
   - Updates automatically as you type

4. **Generate PDF**:
   - Click "Generate PDF" to create your document
   - Choose where to save the file
   - The PDF will include all your information and images in a professional format

## Supported Image Formats

- JPEG (.jpg, .jpeg)
- PNG (.png)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tiff)

## Distribution

### macOS App Distribution:
- The built `.app` file can be copied to `/Applications/`
- Or create a DMG installer for distribution
- The app is properly signed and notarized (if you have developer certificates)

### Windows App Distribution:
- The built `.exe` file is standalone and can be distributed
- No additional dependencies required
- Can be packaged in an installer using tools like Inno Setup

## Customization

The application is built with Python and can be easily customized:

- **Modify the layout**: Edit the `create_widgets()` method
- **Add new fields**: Extend the `fields` list in `create_widgets()`
- **Change PDF styling**: Modify the styles in `create_pdf()`
- **Add new features**: Extend the `PDFBuilderApp` class

## File Structure

```
Property-PDF/
├── pdf_builder_app.py           # Main application file
├── requirements.txt             # Python dependencies
├── PropertyPDFBuilder.spec      # PyInstaller spec for macOS
├── build_macos_app.sh          # macOS build script
├── build_windows_app.bat       # Windows build script
└── README.md                   # This file
```

## Troubleshooting

**Common Issues:**

1. **"Module not found" errors**: Make sure you've installed all requirements with `pip install -r requirements.txt`

2. **Images not loading**: Check that image files are not corrupted and are in supported formats

3. **PDF generation fails**: Ensure you have write permissions in the selected directory

4. **Build fails**: Make sure you have PyInstaller installed and are using the correct build script for your platform

5. **App won't open on macOS**: You may need to right-click and select "Open" the first time, or go to System Preferences > Security & Privacy to allow the app

## Technical Details

- **GUI Framework**: tkinter (built into Python)
- **PDF Generation**: reportlab library
- **Image Processing**: Pillow (PIL)
- **App Bundling**: PyInstaller
- **File Format**: Standard PDF (A4 size)
- **Platform Support**: macOS 10.13+, Windows 10+

## License

This project is open source and available under the MIT License.
