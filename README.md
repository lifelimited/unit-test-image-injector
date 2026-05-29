# Unit Test Image Injector

> Automatically insert device photos into Unit Test Word documents by matching Serial Number folders.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.1.1-orange)

## Overview

This tool automates the process of inserting device photos into Unit Test Word documents (`.docx`). Instead of manually copying and pasting hundreds of images, simply organize your photos into Serial Number folders and let the app do the rest.

### Key Features

- **Auto-match by Serial Number** — Folders named by S/N (e.g., `WZP30149NFC/`) are matched to the corresponding table in the Word document
- **HEIC Support** — iPhone `.heic` photos are automatically converted to JPG
- **Customizable Keyword** — Change the folder prefix from `WZP` to any keyword to match your project
- **Stop Button** — Cancel processing at any time; partial progress is saved
- **Portable** — Standalone `.exe` file, no Python installation required
- **Dark Theme UI** — Modern, clean interface with progress tracking and color-coded logs

## Quick Start

### Option 1: Run the .exe (Recommended)

1. Download `Unit_Test_Image_Injector.exe` from the [Releases](../../releases) page
2. Double-click to run
3. Select your files and click **▶ Run Injection**

### Option 2: Run from Python

```bash
# Install dependencies
pip install python-docx pillow pillow-heif openpyxl

# Run the app
python app.py
```

## How to Use

### Step 1: Prepare Your Folder Structure

Organize device photos into folders named by Serial Number:

```
Photo_Root/
├── WZP30149NFC/
│   └── WZP30149NFC/
│       ├── IMG_8402.JPG
│       ├── IMG_8403.JPG
│       └── 20260526_071125063_iOS.heic    ← auto-converted
├── WZP30149NFE/
│   └── WZP30149NFE/
│       ├── IMG_8410.JPG
│       └── IMG_8411.JPG
└── WZP30149NFJ/
    └── ...
```

### Step 2: Launch the App

![App Screenshot](docs/app_screenshot.png)

### Step 3: Configure

| Field | Description |
|-------|-------------|
| **📄 Word Document** | Select your Unit Test `.docx` template |
| **📁 SN Folders Directory** | Select the parent folder containing all `WZP*` subfolders |
| **Keyword** | Folder prefix to match (default: `WZP`, change as needed) |
| **💾 Output File** | Where to save the result (default: `unit_test_with_images.docx`) |

### Step 4: Run

Click **▶ Run Injection** and watch the progress:

- 🟢 `[OK  ]` — Image inserted successfully
- 🔵 `[SKIP]` — No matching table found in the document
- 🔴 `[ERR ]` — Error inserting image
- 🟡 `[CANCELLED]` — Stopped by user (partial save)

### Step 5: Stop (Optional)

Click **■ Stop** at any time to cancel. Any images already inserted will be saved to the output file.

## Document Structure

The app expects a Word document with **one table per device**, where:

| Cell Location | Content |
|---------------|---------|
| `Row 5, Column D` | Serial Number (e.g., `WZP30149NFC`) |
| `Row 12, Column A` | Empty cell — **images are inserted here** |

> [!NOTE]
> This layout matches the standard Cisco IP Phone Unit Test form used in Thai airport infrastructure projects.

## Supported Image Formats

| Format | Extension | Notes |
|--------|-----------|-------|
| JPEG | `.jpg`, `.jpeg` | Standard |
| PNG | `.png` | Standard |
| BMP | `.bmp` | Standard |
| TIFF | `.tiff`, `.tif` | Standard |
| HEIC | `.heic`, `.heif` | Auto-converted to JPG (iPhone photos) |

## Building from Source

### Build the .exe yourself

```bash
# Install PyInstaller
pip install pyinstaller

# Build standalone executable
python -m PyInstaller --noconfirm --onefile --windowed --name "Unit_Test_Image_Injector" app.py
```

The `.exe` will be created at `dist/Unit_Test_Image_Injector.exe`

## Requirements

### For running from Python:
- Python 3.10+
- `python-docx` — Word document manipulation
- `Pillow` — Image processing
- `pillow-heif` — HEIC/HEIF support
- `openpyxl` — Excel file reading (optional, for MAC scanner)

### For running the .exe:
- Windows 10/11
- No additional dependencies

## Changelog

### v1.1.1
- 🐛 **Fixed empty keyword guard** — Empty keyword no longer matches all folders
- 🐛 **Removed hardcoded auto-detect paths** — User always browses for files (portable)
- 🐛 **Log output optimized** — Large folder sets (600+) show summary instead of flooding log
- ✅ **Multi-keyword support verified** — Works with `WZP`, `FVH`, or any custom prefix

### v1.1.0
- ✅ **Customizable keyword** — Change folder prefix from `WZP` to any string
- ✅ **Stop button** — Cancel processing with partial save
- ✅ **Version display** — Shown in title bar
- ✅ **Improved cancel handling** — Yellow-colored cancel messages in log

### v1.0.0
- ✅ Initial release
- ✅ Serial Number folder matching
- ✅ HEIC auto-conversion
- ✅ Dark theme GUI
- ✅ Progress bar and color-coded log
- ✅ Standalone .exe support

## License

MIT License — feel free to use and modify.
