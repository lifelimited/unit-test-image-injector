#!/bin/bash
cd "$(dirname "$0")"

echo "==========================================="
echo "Setting up Unit Test Image Injector (macOS)"
echo "==========================================="

echo ""
echo "[1/3] Creating Python Virtual Environment (venv)..."
python3 -m venv venv

echo ""
echo "[2/3] Activating virtual environment and installing dependencies..."
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "[3/3] Setup complete!"
echo "To run the application, double-click run_mac.command"
echo ""
read -p "Press Enter to close..."
