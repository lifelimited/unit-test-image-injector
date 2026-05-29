@echo off
echo Installing required libraries (EasyOCR, etc.). This may take a few minutes the first time...
pip install easyocr python-docx pillow-heif
echo.
echo Starting Unit Test Image Injector (Python Script Mode)...
python app.py
pause
