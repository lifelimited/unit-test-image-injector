@echo off
echo ===========================================
echo Setting up Unit Test Image Injector (Windows)
echo ===========================================

echo.
echo [1/3] Creating Python Virtual Environment (venv)...
python -m venv venv

echo.
echo [2/3] Activating virtual environment and installing dependencies...
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo [3/3] Setup complete! 
echo To run the application, double-click run_windows.bat
echo.
pause
