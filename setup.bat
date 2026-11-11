@echo off
rem One-click setup: install Python packages and add the right-click "Convert to" menu.
cd /d "%~dp0"
python --version >nul 2>&1 || (echo Python 3.10+ is required: https://www.python.org/downloads/ & pause & exit /b 1)
echo Installing Python packages...
python -m pip install -r requirements.txt || (echo Package install failed. & pause & exit /b 1)
echo Adding the right-click menu...
python install_menu.py || (echo Menu setup failed. & pause & exit /b 1)
echo.
echo Done. Right-click a file and choose "Convert to".
echo On Windows 11 it is under "Show more options".
pause
