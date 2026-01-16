\
$ErrorActionPreference = "Stop"

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
  Write-Host "[ERROR] Python not found in PATH."
  Write-Host "Install Python from python.org and check 'Add Python to PATH'."
  exit 1
}

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller

if (Test-Path build) { Remove-Item -Recurse -Force build }
if (Test-Path dist) { Remove-Item -Recurse -Force dist }

pyinstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py

Write-Host ""
Write-Host "[INFO] Build complete. Your EXE is in: dist\Copy2.exe"
