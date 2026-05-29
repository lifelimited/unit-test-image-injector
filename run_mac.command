#!/bin/bash
# Move to the folder where the script is located
cd "$(dirname "$0")"

echo "==========================================="
echo "  Unit Test Image Injector (macOS / Linux) "
echo "==========================================="
echo ""
echo "Installing lightweight dependencies (No AI/GPU)..."
pip3 install python-docx Pillow pillow-heif openpyxl

echo ""
echo "Starting Application..."
python3 app.py
