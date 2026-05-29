@echo off
cd /d "%~dp0"

if not exist venv\Scripts\activate (
    echo [Error] Virtual environment not found. 
    echo Please run setup_windows.bat first to install the dependencies!
    pause
    exit /b
)

call venv\Scripts\activate
start pythonw app.py
exit
