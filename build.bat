@echo off
setlocal
cd /d "%~dp0"

REM Build Copy 2.0 Portable EXE (one-file, no console)

where python >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python not found in PATH.
  echo Install Python from python.org and check "Add Python to PATH".
  pause
  exit /b 1
)

python -m pip install --upgrade pip || goto :fail
python -m pip install -r requirements.txt || goto :fail
python -m pip install --upgrade pyinstaller || goto :fail

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Copy2.spec del /q Copy2.spec

set ICON=C:\Users\Ethan\Documents\Programming projects\Copy2Win\assets\Mellowlabs.ico

python -m PyInstaller --noconsole --onefile --name "Copy2" --icon "%ICON%" --add-data "assets;assets" Copy2_Windows.py
python -m PyInstaller --noconsole --onefile --name "Copy2_Uninstall" --icon "%ICON%" --add-data "assets;assets" Copy2_Uninstall.py




echo.
echo [INFO] Build complete. Your EXE is in: %cd%\dist\Copy2.exe
echo.
pause
exit /b 0

:fail
echo.
echo [ERROR] Build failed. Scroll up for the error output.
echo.
pause
exit /b 1
