# ATP Image Injector

> Automatically insert device photos into Unit Test Word documents by matching Serial Number folders.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.3.0-orange)

## Overview

This tool automates the process of inserting device photos into Unit Test Word documents (`.docx`). Instead of manually copying and pasting hundreds of images, simply organize your photos into Serial Number folders and let the app do the rest.

### Key Features

- **Dynamic Table Layout** — Automatically read the Word document so you can visually select where the S/N and Image cells are. No hardcoded formats!
- **Auto-match by Serial Number** — Folders named by S/N (e.g., `WZP301w123/`) are matched to the corresponding table in the Word document
- **HEIC Support** — iPhone `.heic` photos are automatically converted to JPG
- **Customizable Keyword** — Change the folder prefix from `WZP` to any keyword to match your project
- **Stop Button** — Cancel processing at any time; partial progress is saved
- **Portable** — Standalone `.exe` file, no Python installation required
- **Dark Theme UI** — Modern, clean interface with progress tracking and color-coded logs

## Quick Start

### Option 1: Run with Python / GPU (Recommended for v1.3.0)

Since v1.3.0 introduces a powerful AI OCR rotation feature (`easyocr` with PyTorch), compiling it to a standalone `.exe` would create a massive 3.5GB file. We strongly recommend running the app natively through Python to leverage your system's GPU.

1. Double-click `run.bat` (or run `python app.py` from your terminal).
2. The UI will instantly open, ready to use your GPU.

### Option 2: Run the .exe (Legacy / EXIF-only)

1. Download `Unit_Test_Image_Injector_v1.2.1.exe` from the [Releases](../../releases) page
2. Double-click to run (Contains EXIF rotation but NO Smart OCR rotation)
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
├── WZP301000/
│   └── WZP301000/
│       ├── IMG_8402.JPG
│       ├── IMG_8403.JPG
│       └── 20260526_071125063_iOS.heic    ← auto-converted
├── WZP308888/
│   └── WZP308888/
│       ├── IMG_8410.JPG
│       └── IMG_8411.JPG
└── WZP3914ZNNN/
    └── ...
```

### Step 2: Launch the App

![App Screenshot](https://img1.pic.in.th/images/Screenshot-2026-05-29-195205.md.png)

### Step 3: Configure

| Field | Description |
|-------|-------------|
| **📄 Word Document** | Select your Unit Test `.docx` template. **A pop-up will appear** letting you visually select the S/N cell and Image cell. |
| **⚙ Configure Layout** | Click this to re-open the table grid layout and change your selected cell mappings. |
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

## How to Configure Table Layout

The app natively parses the first table in your Word document and displays it in a grid. This allows you to visually map exactly where the Serial Number and Image insertion cells are. No more hardcoded layouts!

![Configure Layout Screenshot](docs/configure_layout.png)

### Step-by-Step Guide:
1. **Select Document:** Click "Browse" next to **Word Document (.docx)**. Once you pick your template, the "Configure Table Layout" pop-up will appear automatically.
2. **Select S/N Cell:** Make sure the **Select S/N Cell (Blue)** option is checked. Click the cell in the grid that contains your Serial Number format (e.g., `WZP301...`). The cell will highlight in **Blue**.
3. **Select Image Cell:** Switch to the **Select Image Cell (Green)** option. Click the empty cell where you want the photos to be injected. The cell will highlight in **Green**.
4. **Approve:** Click the **Approve** button.
5. **Verify:** Back on the main screen, you will see a layout summary (e.g., `Layout: S/N at R6C4 | Image at R13C1`) confirming your configuration.

> **Tip:** If you ever need to change the mapping, you can click the **⚙ Configure Layout** button on the main app screen to reopen the visual grid.

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

### v1.3.0
- 🧠 **Smart Auto-Rotate (EasyOCR)** — Uses PyTorch and your GPU to detect the exact orientation of text on screens and forcefully rotates the image 90°, 180°, or 270° until perfectly readable.

### v1.2.1
- 📸 **EXIF Auto-Rotate** — Automatically normalizes image orientations using EXIF data so they are upright in the document.

### v1.2.0
- ✨ **Dynamic Table Layout** — Automatically reads the Word document so you can visually select where the S/N and Image cells are. No hardcoded formats!
- ♻️ **Major UI Overhaul** — Added pop-up grid preview for selecting exact cell coordinates.

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
