#!/bin/bash
cd "$(dirname "$0")"

if [ ! -f "venv/bin/activate" ]; then
    echo "[Error] Virtual environment not found."
    echo "Please double-click setup_mac.command first to install the dependencies!"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

source venv/bin/activate
python3 app.py
