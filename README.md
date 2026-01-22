# Copy 2.0 (Windows)

Copy 2.0 is a lightweight clipboard history manager for Windows with a simple, fast GUI.  
It continuously tracks your clipboard, lets you browse and search previous copies, pin favorites, and quickly re-copy or combine items.

## Download (recommended)

Use the **Releases** page to download the latest portable ZIP:
- `Copy2.exe` (the app)
- `Copy2_Uninstall.exe` (removes all Copy 2.0 user data and can optionally delete the portable EXEs)

## Run

1. Download the ZIP from **Releases**
2. Extract anywhere (e.g. Desktop)
3. Double-click `Copy2.exe`

## Uninstall / remove all data

Run `Copy2_Uninstall.exe` and confirm the prompts.

**What it removes**
- Your per-user Copy 2.0 data (history/config/favorites) stored under AppData  
  Typical path:
  - `%LOCALAPPDATA%\MellowsLab\copy2\`

It also checks `%APPDATA%` (Roaming) as a fallback.

## Build the EXEs yourself

### Requirements
- Python 3.11/3.12 recommended (from python.org)
- Ensure **Add Python to PATH** is enabled during install

### Build
Open CMD in the repo folder and run:
```bat
build.bat
```

Outputs:
- `dist\Copy2.exe`
- `dist\Copy2_Uninstall.exe`

## Notes
- If Windows SmartScreen appears: **More info → Run anyway**
- This is a portable app: no installer, no admin required

## Discord Link for support and suggestions
- `https://discord.gg/aR7tPND2Jj`

---

## License
Private License
