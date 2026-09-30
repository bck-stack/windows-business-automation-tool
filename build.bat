@echo off
REM Builds dist\URLHealthChecker.exe (single file, no console window)
python -m pip install -r requirements.txt -r requirements-dev.txt || exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --name URLHealthChecker app.py || exit /b 1
echo.
echo Built dist\URLHealthChecker.exe
