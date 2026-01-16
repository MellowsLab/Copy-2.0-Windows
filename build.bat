@echo off
setlocal

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

python -m PyInstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py || goto :fail

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
