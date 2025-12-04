# QR Code Generator for Revit 2023

A Revit add-in that generates and inserts QR codes into drawing sheets for Autodesk Revit. Built for improving construction documentation workflows.

## What It Does

This add-in helps embed construction data into QR codes that can be placed on Revit sheets. Each QR code contains:

- Drawing Number
- Sheet Name
- Revision
- Issue Date
- Checked By

## Features

- **Two ways to work**: Manually enter data or automatically pull it from your current sheet
- **Live preview**: See your QR code before inserting it
- **Automatic sizing**: QR codes are inserted at 2" x 2", ready to use
- **Save option**: Export QR codes as PNG files for other uses

## Installation

1. Download the latest release
2. Copy `QRCoder.dll`, `QRCodeRevitAddin.dll` and `QRCodeRevitAddin.addin` to:
   ```
   %APPDATA%\Autodesk\Revit\Addins\2023\
   ```
3. Restart Revit
4. Look for the "QR Tools" tab in your ribbon

## How to Use

### Quick Insert from Sheet
1. Open a sheet view in Revit
2. Click **Quick Insert from Sheet** in the QR Tools ribbon
3. The window opens with your sheet data already filled in
4. Click **Generate QR Code**
5. Click **Insert into Sheet**

### Manual Entry
1. Click **Generate QR Code** in the QR Tools ribbon
2. Fill in your document information
3. Click **Generate QR Code** to preview
4. Click **Insert into Sheet** when ready

## Requirements

- Revit 2023
- .NET Framework 4.8
- Windows (x64)

## Development

- **Commands**: Handle Revit interaction
- **Domain**: Operational logic for QR generation and sheet manipulation
- **Models**: Data structures
- **ViewModels**: UI logic
- **Views**: WPF windows

## License

Property of DEAXO GmbH Dresden
Copyright © 2025. All rights reserved.

## Support

For issues or questions, please open an issue on GitHub.

---

*Developed with love by DEAXO BIM-Team.*
