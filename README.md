\
# Copy 2.0 (Windows Portable)

This ZIP contains a Windows-portable version of Copy 2.0 that you can package into a **single .exe** (no installer).

## Build a single EXE
1. Install Python (3.11 or 3.12 recommended) from python.org and ensure **Add Python to PATH** is checked.
2. Open CMD or PowerShell in this folder.
3. Run:

### CMD
```bat
build.bat
```

### PowerShell
```powershell
.\build.ps1
```

Output:
- `dist\Copy2.exe`

## Run from source (optional)
```powershell
python .\Copy2_Windows.py
```

## Shortcuts (when app is focused)
- Ctrl+F: focus search
- Enter: search
- Ctrl+C: copy selected
- Del: delete selected
- Ctrl+E: export
- Ctrl+I: import
- Ctrl+L: clear all
- Esc: clear search
