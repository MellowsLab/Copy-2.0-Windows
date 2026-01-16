"""
Copy 2.0 Uninstaller (Windows Portable)

- Removes Copy 2.0 data stored under the current user's AppData folders.
- Optionally deletes Copy2.exe and the uninstaller itself from the current folder.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

from platformdirs import user_data_dir

APP_ID = "copy2"
VENDOR = "MellowsLab"


def candidate_dirs() -> list[Path]:
    dirs: list[Path] = []

    # Canonical platformdirs locations
    try:
        dirs.append(Path(user_data_dir(APP_ID, VENDOR)))                 # usually %LOCALAPPDATA%\MellowsLab\copy2
    except Exception:
        pass
    try:
        dirs.append(Path(user_data_dir(APP_ID, VENDOR, roaming=True)))   # %APPDATA%\MellowsLab\copy2 (rare)
    except Exception:
        pass

    # Fallback/legacy guesses
    appdata = os.environ.get("APPDATA")
    localapp = os.environ.get("LOCALAPPDATA")
    if appdata:
        dirs.extend([Path(appdata) / VENDOR / APP_ID, Path(appdata) / APP_ID])
    if localapp:
        dirs.extend([Path(localapp) / VENDOR / APP_ID, Path(localapp) / APP_ID])

    # Deduplicate
    seen = set()
    out = []
    for d in dirs:
        key = str(d).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out


def remove_dir(p: Path) -> bool:
    if not p.exists():
        return True
    try:
        shutil.rmtree(p, ignore_errors=False)
        return True
    except Exception:
        try:
            shutil.rmtree(p, ignore_errors=True)
            return not p.exists()
        except Exception:
            return False


def schedule_delete(files: list[Path]) -> None:
    # Windows can't delete a running exe; schedule via temp batch.
    bat = Path(tempfile.gettempdir()) / "copy2_uninstall_cleanup.bat"
    targets = " ".join(f'"{str(p)}"' for p in files)
    bat.write_text(
        f"""@echo off
timeout /t 2 /nobreak >nul
for %%F in ({targets}) do (
  del /f /q "%%~F" >nul 2>&1
)
del /f /q "%~f0" >nul 2>&1
""",
        encoding="utf-8",
    )
    os.startfile(str(bat))


def main():
    root = tk.Tk()
    root.withdraw()

    dirs = candidate_dirs()
    existing = [d for d in dirs if d.exists()]

    msg = "This will remove Copy 2.0 data stored on this Windows user profile.\n\n"
    if existing:
        msg += "Folders to remove:\n" + "\n".join(f" - {d}" for d in existing) + "\n\n"
    else:
        msg += "No Copy 2.0 data folders were found.\n\n"
    msg += "Proceed with uninstall?"

    if not messagebox.askyesno("Copy 2.0 Uninstall", msg):
        root.destroy()
        return

    failures = []
    for d in existing:
        if not remove_dir(d):
            failures.append(str(d))

    # Ask about deleting portable EXEs from current folder
    if messagebox.askyesno("Copy 2.0 Uninstall", "Also delete Copy2.exe and this uninstaller from this folder?"):
        here = Path(sys.argv[0]).resolve()
        folder = here.parent
        copy2_exe = folder / "Copy2.exe"
        to_delete = []
        if copy2_exe.exists():
            to_delete.append(copy2_exe)
        if here.exists():
            to_delete.append(here)
        if to_delete:
            schedule_delete(to_delete)

    if failures:
        messagebox.showwarning("Copy 2.0 Uninstall", "Some folders could not be removed:\n" + "\n".join(failures))
    else:
        messagebox.showinfo("Copy 2.0 Uninstall", "Cleanup complete.")

    root.destroy()


if __name__ == "__main__":
    main()
