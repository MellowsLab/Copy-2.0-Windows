"""
Copy 2.0 (Windows) — Modern UI Edition (Improved + Favorites Safe + Self-Updating ZIP Releases)

Key fixes/changes:
- Favorites are protected:
  - Clean keeps favorites (★) and removes non-favorites only.
  - Favorites are never evicted when reaching max_history.
  - When history is full and only favorites remain, new items are blocked and user is prompted to increase max or remove favorites.
- Capacity warnings:
  - When you reach your configured max_history: you get a one-time session prompt to consider increasing.
  - When you reach the hard cap (500): you get a one-time session warning about allocated limit.
- Controls/Hotkeys panel added to Settings.
- Import/Export:
  - Exports settings/history/favorites.
  - Import can optionally apply settings, merges history, preserves favorites.

Auto-update (matches your GitHub Releases pattern):
- On launch: pings GitHub "latest release" for MellowsLab/Copy-2.0-Windows.
- If a newer version exists: prompt user.
  - Cancel = do nothing
  - Update = download latest ZIP asset (e.g., Copy2.zip), extract Copy2.exe (+ Copy2_Uninstall.exe if present),
            then swap files via an updater .bat after the app exits, then restart.
- Also includes "Check Updates" button + Ctrl+U hotkey.

Dependencies:
  pyperclip
  platformdirs
Recommended:
  ttkbootstrap

Build:
  python -m PyInstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import tempfile
import threading
import subprocess
import webbrowser
import urllib.request
import zipfile
from collections import deque
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pyperclip
from platformdirs import user_data_dir

# -----------------------------
# Version + GitHub Update Config
# -----------------------------
# IMPORTANT: Set this to the version of THIS build you are distributing.
# Example: if the exe you ship is v1.0.1, set APP_VERSION = "1.0.1"
APP_VERSION = "1.0.1"

# Your repo (from your link)
GITHUB_OWNER = "MellowsLab"
GITHUB_REPO = "Copy-2.0-Windows"

# Your releases use a ZIP asset (screenshot shows Copy2.zip)
GITHUB_ASSET_NAME_HINT = "Copy2"  # prefers assets containing this substring (case-insensitive)
AUTO_CHECK_UPDATES_ON_LAUNCH = True

# -----------------------------
# Optional modern theme layer
# -----------------------------
USE_TTKB = True
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
except Exception:
    USE_TTKB = False
    from tkinter import ttk  # type: ignore

    PRIMARY = "primary"
    SECONDARY = "secondary"
    SUCCESS = "success"
    INFO = "info"
    WARNING = "warning"
    DANGER = "danger"
    OUTLINE = "outline"
    BOTH = tk.BOTH
    X = tk.X
    Y = tk.Y
    LEFT = tk.LEFT
    RIGHT = tk.RIGHT
    TOP = tk.TOP
    BOTTOM = tk.BOTTOM


APP_NAME = "Copy 2.0"
APP_ID = "copy2"
VENDOR = "MellowsLab"

DEFAULT_MAX_HISTORY = 50
DEFAULT_POLL_MS = 400
HARD_MAX_HISTORY = 500


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_json_load(path: Path, default):
    try:
        if not path.exists():
            return default
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def safe_json_save(path: Path, obj) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


class ToolTip:
    """Simple hover tooltip for Tk widgets."""
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show, add=True)
        widget.bind("<Leave>", self._hide, add=True)

    def _show(self, _event=None):
        try:
            if self.tip or not self.text:
                return
            x = self.widget.winfo_rootx() + 15
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
            self.tip = tk.Toplevel(self.widget)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{x}+{y}")
            lbl = tk.Label(
                self.tip,
                text=self.text,
                justify="left",
                background="#111",
                foreground="#fff",
                relief="solid",
                borderwidth=1,
                font=("Segoe UI", 9),
            )
            lbl.pack(ipadx=8, ipady=4)
        except Exception:
            self.tip = None

    def _hide(self, _event=None):
        try:
            if self.tip:
                self.tip.destroy()
        except Exception:
            pass
        self.tip = None


@dataclass
class Settings:
    max_history: int = DEFAULT_MAX_HISTORY
    poll_ms: int = DEFAULT_POLL_MS
    session_only: bool = False
    reverse_lines_copy: bool = False
    theme: str = "flatly"

    @staticmethod
    def from_dict(d: dict) -> "Settings":
        s = Settings()
        s.max_history = int(d.get("max_history", DEFAULT_MAX_HISTORY))
        s.poll_ms = int(d.get("poll_ms", DEFAULT_POLL_MS))
        s.session_only = bool(d.get("session_only", False))
        s.reverse_lines_copy = bool(d.get("reverse_lines_copy", False))
        s.theme = str(d.get("theme", "flatly"))
        s.max_history = max(5, min(HARD_MAX_HISTORY, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        return s


class Copy2AppBase:
    """Shared logic for both ttkbootstrap and ttk fallback variants."""

    def _init_state(self):
        self.data_dir = Path(user_data_dir(APP_ID, VENDOR))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"

        self.settings = Settings.from_dict(safe_json_load(self.settings_path, {}))

        loaded_history = safe_json_load(self.history_path, [])
        loaded_favs = safe_json_load(self.favs_path, [])

        # Normalize favorites
        self.favorites = []
        seen_f = set()
        if isinstance(loaded_favs, list):
            for x in loaded_favs:
                if isinstance(x, str) and x not in seen_f:
                    self.favorites.append(x)
                    seen_f.add(x)

        # History WITHOUT deque maxlen (we enforce capacity ourselves to protect favorites)
        base_items = []
        if isinstance(loaded_history, list):
            for x in loaded_history:
                if isinstance(x, str) and x.strip():
                    base_items.append(x)

        # Ensure favorites always exist in history
        for fav in self.favorites:
            if fav not in base_items:
                base_items.append(fav)

        self.history = deque(base_items)

        self.paused = False
        self.last_clip = ""

        # Search
        self.search_matches: list[str] = []
        self.search_index = 0
        self.search_query = ""

        # Polling
        self._poll_job = None

        # View model
        self.view_items: list[str] = []

        # Selection ordering for Combine
        self._prev_sel_set: set[int] = set()
        self._sel_order: list[int] = []

        # Preview edit model
        self._selected_item_text: str | None = None
        self._preview_dirty = False

        # Capacity warnings (one-time per session)
        self._warned_soft_cap = False
        self._warned_hard_cap = False
        self._warned_block_add = False

        # Tooltips storage to avoid GC
        self._tooltips: list[ToolTip] = []

        # Update check state
        self._update_check_inflight = False
        self._last_update_info = None

        # Enforce capacity now
        self._ensure_capacity_for_favorites(notify=False)
        self._apply_capacity_preserving_favorites(reason="startup")

    def _persist(self):
        if self.settings.session_only:
            return
        safe_json_save(self.settings_path, asdict(self.settings))
        safe_json_save(self.history_path, list(self.history))
        safe_json_save(self.favs_path, list(self.favorites))

    # -----------------------------
    # Tooltip helper
    # -----------------------------
    def _tip(self, widget, text: str):
        try:
            self._tooltips.append(ToolTip(widget, text))
        except Exception:
            pass

    # -----------------------------
    # Update helpers (GitHub latest + ZIP asset)
    # -----------------------------
    def _is_frozen(self) -> bool:
        return bool(getattr(sys, "frozen", False)) and hasattr(sys, "executable")

    def _github_api_latest_url(self) -> str:
        return f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

    def _compare_versions(self, a: str, b: str) -> int:
        """-1 if a<b, 0 if equal, 1 if a>b. Handles tags like v1.0.1."""
        def norm(v: str):
            v = (v or "").strip()
            if v.lower().startswith("v"):
                v = v[1:]
            parts = []
            for p in v.split("."):
                m = re.match(r"(\d+)", p.strip())
                parts.append(int(m.group(1)) if m else 0)
            return parts
        A, B = norm(a), norm(b)
        n = max(len(A), len(B))
        A += [0] * (n - len(A))
        B += [0] * (n - len(B))
        return (A > B) - (A < B)

    def _fetch_latest_release_info(self) -> dict | None:
        """
        Returns:
          version, html_url, notes, asset_name, asset_url, asset_size
        Prefers ZIP assets (your releases ship Copy2.zip).
        """
        try:
            req = urllib.request.Request(
                self._github_api_latest_url(),
                headers={
                    "User-Agent": f"{APP_NAME}/{APP_VERSION}",
                    "Accept": "application/vnd.github+json",
                },
                method="GET",
            )
            with urllib.request.urlopen(req, timeout=6) as r:
                raw = r.read().decode("utf-8", errors="replace")
            data = json.loads(raw)

            tag = str(data.get("tag_name", "")).strip()
            html_url = str(data.get("html_url", "")).strip()
            notes = str(data.get("body", "")).strip()

            assets = data.get("assets", []) if isinstance(data.get("assets", []), list) else []
            zip_assets = []
            for a in assets:
                if not isinstance(a, dict):
                    continue
                name = str(a.get("name", "")).strip()
                dl = str(a.get("browser_download_url", "")).strip()
                size = int(a.get("size", 0) or 0)
                if name.lower().endswith(".zip") and dl:
                    zip_assets.append({"name": name, "url": dl, "size": size})

            chosen = None
            if zip_assets:
                if GITHUB_ASSET_NAME_HINT:
                    hint = GITHUB_ASSET_NAME_HINT.lower()
                    for a in zip_assets:
                        if hint in a["name"].lower():
                            chosen = a
                            break
                if chosen is None:
                    chosen = zip_assets[0]

            info = {
                "version": tag,
                "html_url": html_url,
                "notes": notes,
                "asset_name": chosen["name"] if chosen else "",
                "asset_url": chosen["url"] if chosen else "",
                "asset_size": int(chosen["size"]) if chosen else 0,
            }
            if not info["version"]:
                return None
            return info
        except Exception:
            return None

    def _check_updates_async(self, prompt_if_new: bool = True):
        if self._update_check_inflight:
            return
        self._update_check_inflight = True

        def worker():
            info = self._fetch_latest_release_info()

            def done():
                self._update_check_inflight = False
                self._last_update_info = info

                if info is None:
                    if not prompt_if_new:
                        messagebox.showwarning(APP_NAME, "Could not check for updates.")
                    self.status_var.set(f"Update check failed — {now_ts()}")
                    return

                latest = str(info.get("version", "")).strip()
                if self._compare_versions(APP_VERSION, latest) >= 0:
                    if not prompt_if_new:
                        messagebox.showinfo(APP_NAME, f"You're up to date.\n\nInstalled: {APP_VERSION}\nLatest: {latest}")
                    else:
                        self.status_var.set(f"Up to date (Latest: {latest}) — {now_ts()}")
                    return

                # New version
                if prompt_if_new:
                    self._prompt_update(info)
                else:
                    self._prompt_update(info)

            try:
                self.after(0, done)
            except Exception:
                self._update_check_inflight = False

        threading.Thread(target=worker, daemon=True).start()

    def _prompt_update(self, info: dict):
        latest = str(info.get("version", "")).strip()
        html_url = str(info.get("html_url", "")).strip()
        notes = str(info.get("notes", "")).strip()
        asset_url = str(info.get("asset_url", "")).strip()
        asset_name = str(info.get("asset_name", "")).strip()

        if not asset_url:
            msg = (
                f"Update available.\n\nInstalled: {APP_VERSION}\nLatest: {latest}\n\n"
                "No ZIP asset was found on the latest release.\nOpen the release page?"
            )
            if messagebox.askyesno(APP_NAME, msg):
                if html_url:
                    webbrowser.open(html_url)
            return

        notes_trim = notes.strip()
        if len(notes_trim) > 600:
            notes_trim = notes_trim[:600] + "..."

        msg = (
            f"Update available.\n\n"
            f"Installed: {APP_VERSION}\nLatest: {latest}\n\n"
            f"Download: {asset_name or 'Release ZIP'}\n"
        )
        if notes_trim:
            msg += f"\nRelease notes:\n{notes_trim}\n"
        msg += "\nInstall this update now?\n\nYes = Update + Restart\nNo = Cancel"

        if not messagebox.askyesno(APP_NAME, msg):
            self.status_var.set(f"Update skipped (Latest: {latest}) — {now_ts()}")
            return

        self._download_and_apply_update_async(info)

    def _download_and_apply_update_async(self, info: dict):
        asset_url = str(info.get("asset_url", "")).strip()
        latest = str(info.get("version", "")).strip()

        if not self._is_frozen():
            messagebox.showinfo(
                APP_NAME,
                "Auto-update is supported for the packaged portable .exe.\n\n"
                "You are running a Python script environment; opening the release page instead."
            )
            html_url = str(info.get("html_url", "")).strip()
            if html_url:
                webbrowser.open(html_url)
            return

        target_exe = Path(sys.executable)
        target_dir = target_exe.parent

        if not target_exe.exists():
            messagebox.showerror(APP_NAME, "Could not locate the current executable for updating.")
            return

        if not os.access(str(target_dir), os.W_OK):
            messagebox.showwarning(
                APP_NAME,
                "The app does not have permission to write to its folder.\n\n"
                "Move the app to a writable location (e.g., your Desktop) or run as Administrator."
            )
            return

        self.status_var.set(f"Downloading update {latest}... — {now_ts()}")

        def worker():
            tmp_dir = Path(tempfile.mkdtemp(prefix="copy2_update_"))
            zip_path = tmp_dir / f"Copy2_update_{latest}.zip"
            extract_dir = tmp_dir / "extracted"
            err = None

            try:
                req = urllib.request.Request(
                    asset_url,
                    headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"},
                    method="GET",
                )
                with urllib.request.urlopen(req, timeout=20) as r:
                    with open(zip_path, "wb") as f:
                        while True:
                            chunk = r.read(1024 * 256)
                            if not chunk:
                                break
                            f.write(chunk)

                if not zip_path.exists() or zip_path.stat().st_size < 1024 * 10:
                    raise RuntimeError("Downloaded ZIP is missing or unexpectedly small.")

                extract_dir.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(zip_path, "r") as z:
                    z.extractall(extract_dir)

                # Find expected files inside zip
                new_copy2 = None
                new_uninst = None
                for p in extract_dir.rglob("*"):
                    if not p.is_file():
                        continue
                    name = p.name.lower()
                    if name == "copy2.exe":
                        new_copy2 = p
                    elif name == "copy2_uninstall.exe":
                        new_uninst = p

                if new_copy2 is None:
                    raise RuntimeError("Copy2.exe not found inside the downloaded ZIP.")

                # Stage files with exact names for the updater script
                staged_copy2 = tmp_dir / "Copy2.exe"
                staged_uninst = tmp_dir / "Copy2_Uninstall.exe"

                # Copy bytes (avoid cross-volume move edge cases)
                staged_copy2.write_bytes(new_copy2.read_bytes())
                if new_uninst is not None:
                    staged_uninst.write_bytes(new_uninst.read_bytes())

            except Exception as e:
                err = str(e)

            def done():
                if err:
                    self.status_var.set(f"Update download failed — {now_ts()}")
                    messagebox.showwarning(APP_NAME, f"Update failed:\n\n{err}")
                    return

                ok = self._launch_updater_and_exit(
                    staged_copy2=tmp_dir / "Copy2.exe",
                    staged_uninst=(tmp_dir / "Copy2_Uninstall.exe"),
                    target_exe=target_exe,
                    target_dir=target_dir,
                )
                if not ok:
                    messagebox.showwarning(APP_NAME, "Failed to start the updater process. Update was not applied.")

            try:
                self.after(0, done)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _launch_updater_and_exit(self, staged_copy2: Path, staged_uninst: Path, target_exe: Path, target_dir: Path) -> bool:
        """
        Uses a .bat in the staging directory:
        - waits for current PID to exit
        - moves staged Copy2.exe over the running exe path
        - moves staged Copy2_Uninstall.exe into the same folder (if present)
        - restarts the app
        """
        try:
            pid = os.getpid()
            bat = staged_copy2.parent / "copy2_updater.bat"

            # In case the running exe isn't named Copy2.exe, still replace the exact sys.executable path.
            bat_contents = f"""@echo off
setlocal enabledelayedexpansion
set PID={pid}
set TARGET_EXE="{str(target_exe)}"
set TARGET_DIR="{str(target_dir)}"
set NEW_COPY2="{str(staged_copy2)}"
set NEW_UNINST="{str(staged_uninst)}"

:wait
for /f "tokens=2 delims=," %%A in ('tasklist /FI "PID eq %PID%" /FO CSV /NH 2^>NUL') do (
  if "%%~A"=="%PID%" (
    timeout /t 1 /nobreak >NUL
    goto wait
  )
)

REM Replace main exe (overwrite)
move /y %NEW_COPY2% %TARGET_EXE% >NUL 2>&1
if errorlevel 1 (
  timeout /t 1 /nobreak >NUL
  move /y %NEW_COPY2% %TARGET_EXE% >NUL 2>&1
)

REM Replace optional uninstaller if staged
if exist %NEW_UNINST% (
  move /y %NEW_UNINST% %TARGET_DIR%\\Copy2_Uninstall.exe >NUL 2>&1
)

start "" %TARGET_EXE%
del "%~f0"
endlocal
"""
            bat.write_text(bat_contents, encoding="utf-8")

            creationflags = 0
            if hasattr(subprocess, "DETACHED_PROCESS"):
                creationflags |= subprocess.DETACHED_PROCESS
            if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
                creationflags |= subprocess.CREATE_NEW_PROCESS_GROUP

            subprocess.Popen(
                ["cmd.exe", "/c", str(bat)],
                close_fds=True,
                creationflags=creationflags,
                cwd=str(staged_copy2.parent),
            )

            self.status_var.set(f"Applying update and restarting... — {now_ts()}")
            self.after(150, self._on_close_for_update)
            return True
        except Exception:
            return False

    def _on_close_for_update(self):
        # Close without prompting about unsaved preview edits (user accepted restart)
        try:
            if self._poll_job is not None:
                self.after_cancel(self._poll_job)
        except Exception:
            pass
        try:
            self._persist()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    # -----------------------------
    # Capacity helpers (favorites-safe)
    # -----------------------------
    def _ensure_capacity_for_favorites(self, notify: bool = True):
        fav_count = len(self.favorites)
        if fav_count <= self.settings.max_history:
            return

        if fav_count <= HARD_MAX_HISTORY:
            old = self.settings.max_history
            self.settings.max_history = fav_count
            self._persist()
            if notify:
                messagebox.showinfo(
                    APP_NAME,
                    f"Favorites count ({fav_count}) exceeded Max history ({old}).\n\n"
                    f"Max history increased to {fav_count} to protect favorites."
                )
        else:
            if notify and not self._warned_hard_cap:
                messagebox.showwarning(
                    APP_NAME,
                    f"Favorites exceed the hard cap ({HARD_MAX_HISTORY}).\n\n"
                    "You have used all allocated memory under the hard-coded limit.\n"
                    "Please remove some favorites."
                )
                self._warned_hard_cap = True

    def _apply_capacity_preserving_favorites(self, reason: str = "") -> bool:
        items = list(self.history)
        fav_set = set(self.favorites)

        if len(items) <= self.settings.max_history:
            self.history = deque(items)
            self._maybe_warn_caps(len(items))
            return True

        while len(items) > self.settings.max_history:
            removed_one = False
            for idx, val in enumerate(items):
                if val not in fav_set:
                    items.pop(idx)
                    removed_one = True
                    break
            if not removed_one:
                self.history = deque(items)
                self._maybe_warn_caps(len(items))
                return False

        self.history = deque(items)
        self._maybe_warn_caps(len(items))
        return True

    def _maybe_warn_caps(self, current_len: int):
        if current_len >= self.settings.max_history and not self._warned_soft_cap:
            try:
                messagebox.showinfo(
                    APP_NAME,
                    f"You have reached Max history ({self.settings.max_history}).\n\n"
                    "New items will remove older non-favorites.\n"
                    "Consider increasing Max history in Settings."
                )
            except Exception:
                pass
            self._warned_soft_cap = True

        if self.settings.max_history >= HARD_MAX_HISTORY and current_len >= HARD_MAX_HISTORY and not self._warned_hard_cap:
            try:
                messagebox.showwarning(
                    APP_NAME,
                    f"You have reached the hard-coded memory limit ({HARD_MAX_HISTORY}).\n\n"
                    "You have used up all allocated memory under the hard-coded limit.\n"
                    "Please consider removing old favorites or cleaning up."
                )
            except Exception:
                pass
            self._warned_hard_cap = True

    # -----------------------------
    # Clipboard polling
    # -----------------------------
    def _poll_clipboard(self):
        try:
            if not self.paused:
                text = pyperclip.paste()
                if isinstance(text, str):
                    text = text.strip("\r\n")
                else:
                    text = ""
                if text and text != self.last_clip:
                    self.last_clip = text
                    self._add_history_item(text)
        except Exception:
            pass

        self._poll_job = self.after(self.settings.poll_ms, self._poll_clipboard)

    def _add_history_item(self, text: str):
        items = list(self.history)

        if items and items[-1] == text:
            return

        if text in items:
            items = [x for x in items if x != text]
        items.append(text)

        self.history = deque(items)
        ok = self._apply_capacity_preserving_favorites(reason="add_item")

        if not ok:
            # Roll back: remove the new item
            items = [x for x in items if x != text]
            self.history = deque(items)
            self._apply_capacity_preserving_favorites(reason="rollback")

            if not self._warned_block_add:
                messagebox.showwarning(
                    APP_NAME,
                    f"Cannot add new clipboard item.\n\n"
                    f"History is full ({self.settings.max_history}) and remaining items are protected favorites.\n"
                    "Increase Max history in Settings, or remove some favorites."
                )
                self._warned_block_add = True

            self.status_var.set(f"History full (favorites protected). Increase Max history. — {now_ts()}")
            return

        self._refresh_list(select_last=True)
        self._persist()

    # -----------------------------
    # List/view helpers
    # -----------------------------
    def _format_list_item(self, item: str) -> str:
        one = re.sub(r"\s+", " ", item).strip()
        if len(one) > 90:
            one = one[:87] + "..."
        prefix = "★ " if item in self.favorites else "  "
        return prefix + one

    def _current_filter(self) -> str:
        if hasattr(self, "filter_var") and self.filter_var is not None:
            try:
                return str(self.filter_var.get())
            except Exception:
                return "all"
        return "all"

    def _refresh_list(self, select_last: bool = False):
        items = list(self.history)
        if self._current_filter() == "fav":
            items = [x for x in items if x in self.favorites]

        self.view_items = items

        self.listbox.delete(0, tk.END)
        for item in self.view_items:
            self.listbox.insert(tk.END, self._format_list_item(item))

        self._prev_sel_set = set()
        self._sel_order = []

        if select_last and self.view_items:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(tk.END)
            self.listbox.see(tk.END)
            self._on_select()

        self.status_var.set(
            f"Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}   v{APP_VERSION}"
        )

    def _get_selected_indices(self) -> list[int]:
        return list(self.listbox.curselection())

    def _get_selected_text(self) -> str | None:
        sel = self._get_selected_indices()
        if not sel:
            return None
        i = sel[0]
        if i < 0 or i >= len(self.view_items):
            return None
        return self.view_items[i]

    def _apply_reverse_lines(self, text: str) -> str:
        lines = text.splitlines()
        lines.reverse()
        return "\n".join(lines)

    def _get_preview_display_text_for_item(self, item_text: str) -> str:
        if self.settings.reverse_lines_copy:
            return self._apply_reverse_lines(item_text)
        return item_text

    def _set_preview_text(self, text: str, mark_clean: bool = True):
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", text)
        self._preview_dirty = not mark_clean
        self._update_preview_dirty_ui()

        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _on_select(self):
        if self._preview_dirty and self._selected_item_text is not None:
            resp = messagebox.askyesnocancel(
                APP_NAME,
                "You have unsaved edits in Preview.\n\n"
                "Yes = Save edits\nNo = Discard edits\nCancel = Stay on current item"
            )
            if resp is None:
                self._reselect_current_item()
                return
            if resp is True:
                self._save_preview_edits()
            else:
                self._revert_preview_edits()

        t = self._get_selected_text()
        if not t:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)
            return

        self._selected_item_text = t
        display = self._get_preview_display_text_for_item(t)
        self._set_preview_text(display, mark_clean=True)

        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _reselect_current_item(self):
        if not self._selected_item_text:
            return
        try:
            if self._selected_item_text not in self.view_items:
                return
            idx = self.view_items.index(self._selected_item_text)
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        except Exception:
            pass

    # -----------------------------
    # Selection order tracking for Combine
    # -----------------------------
    def _on_listbox_select_event(self, _event=None):
        current = set(self.listbox.curselection())
        added = current - self._prev_sel_set
        removed = self._prev_sel_set - current

        if removed:
            self._sel_order = [i for i in self._sel_order if i not in removed]

        if added:
            for i in sorted(list(added)):
                if i not in self._sel_order:
                    self._sel_order.append(i)

        self._prev_sel_set = current

        if len(current) == 1:
            self._on_select()

    # -----------------------------
    # Actions
    # -----------------------------
    def _toggle_pause(self):
        self.paused = bool(self.pause_var.get())
        if hasattr(self, "capture_state_var"):
            self.capture_state_var.set("Paused" if self.paused else "Capturing")
        self.status_var.set(f"{'Paused' if self.paused else 'Capturing'} — {now_ts()}")

    def _toggle_reverse_lines(self):
        self.settings.reverse_lines_copy = bool(self.reverse_var.get())
        self._persist()

        if self._selected_item_text is not None:
            display = self._get_preview_display_text_for_item(self._selected_item_text)
            if self._preview_dirty:
                resp = messagebox.askyesno(APP_NAME, "Reverse-lines will re-render Preview.\nDiscard current unsaved edits?")
                if not resp:
                    return
                self._preview_dirty = False
            self._set_preview_text(display, mark_clean=True)

    def _update_preview_dirty_ui(self):
        if hasattr(self, "preview_dirty_var") and self.preview_dirty_var is not None:
            self.preview_dirty_var.set("Edited" if self._preview_dirty else "")
        if hasattr(self, "save_btn") and self.save_btn is not None:
            try:
                self.save_btn.configure(state=("normal" if self._preview_dirty else "disabled"))
            except Exception:
                pass
        if hasattr(self, "revert_btn") and self.revert_btn is not None:
            try:
                self.revert_btn.configure(state=("normal" if self._preview_dirty else "disabled"))
            except Exception:
                pass

    def _mark_preview_dirty(self, _event=None):
        if self._selected_item_text is None and self.preview.get("1.0", tk.END).strip():
            self._preview_dirty = True
        elif self._selected_item_text is not None:
            baseline = self._get_preview_display_text_for_item(self._selected_item_text)
            current = self.preview.get("1.0", tk.END).rstrip("\n")
            self._preview_dirty = (current != baseline)
        self._update_preview_dirty_ui()

        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _save_preview_edits(self, _event=None):
        text_now = self.preview.get("1.0", tk.END).rstrip("\n")

        if self._selected_item_text is None:
            if text_now.strip():
                self._add_history_item(text_now)
                self.status_var.set(f"Saved new item from Preview — {now_ts()}")
                self._preview_dirty = False
                self._update_preview_dirty_ui()
            return

        old = self._selected_item_text
        new = text_now

        if not new.strip():
            messagebox.showwarning(APP_NAME, "Cannot save an empty item.")
            return

        items = [x for x in self.history if x != old]
        items.append(new)
        self.history = deque(items)

        if old in self.favorites:
            self.favorites = [new if x == old else x for x in self.favorites]

        self._ensure_capacity_for_favorites(notify=False)
        ok = self._apply_capacity_preserving_favorites(reason="save_edit")
        if not ok:
            messagebox.showwarning(APP_NAME, "Saved, but capacity constraints require increasing Max history.")

        self._selected_item_text = new
        self._preview_dirty = False
        self._persist()
        self._refresh_list(select_last=True)
        self.status_var.set(f"Saved edits — {now_ts()}")

    def _revert_preview_edits(self, _event=None):
        if self._selected_item_text is None:
            self._set_preview_text("", mark_clean=True)
            return
        display = self._get_preview_display_text_for_item(self._selected_item_text)
        self._set_preview_text(display, mark_clean=True)
        self.status_var.set(f"Reverted edits — {now_ts()}")

    def _copy_selected(self):
        out = self.preview.get("1.0", tk.END).rstrip("\n")
        if not out.strip():
            return
        try:
            pyperclip.copy(out)
            self.status_var.set(f"Copied — {now_ts()}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to copy to clipboard:\n{e}")

    def _delete_selected(self):
        t = self._get_selected_text()
        if not t:
            return

        items = [x for x in self.history if x != t]
        self.history = deque(items)

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]

        if self._selected_item_text == t:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)

        self._apply_capacity_preserving_favorites(reason="delete")
        self._refresh_list(select_last=True)
        self._persist()

    def _clean_history_keep_favorites(self):
        fav_set = set(self.favorites)
        kept = [x for x in self.history if x in fav_set]
        self.history = deque(kept)

        if self._selected_item_text and self._selected_item_text not in kept:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)

        self._refresh_list(select_last=True)
        self._persist()
        self.status_var.set(f"Cleaned history (favorites kept) — {now_ts()}")

    def _clear_all_history(self):
        resp = messagebox.askyesnocancel(
            APP_NAME,
            "This will remove ALL history items, including Favorites.\n\n"
            "Yes = Clear all\nNo/Cancel = Abort"
        )
        if resp is not True:
            return
        self.history = deque([])
        self.favorites = []
        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._refresh_list()
        self._persist()
        self.status_var.set(f"Cleared all history — {now_ts()}")

    def _clear_history(self):
        if not messagebox.askyesno(APP_NAME, "Clean history by removing non-favorites (keep ★ items)?"):
            return
        self._clean_history_keep_favorites()

    def _toggle_favorite_selected(self):
        t = self._get_selected_text()
        if not t:
            return

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]
        else:
            self.favorites.append(t)

        seen = set()
        out = []
        for x in self.favorites:
            if x not in seen:
                out.append(x)
                seen.add(x)
        self.favorites = out

        self._ensure_capacity_for_favorites(notify=True)

        if t not in self.history:
            self.history.append(t)
            self._apply_capacity_preserving_favorites(reason="fav_add")

        self._refresh_list()
        self._persist()

    def _combine_selected(self):
        sel = self._get_selected_indices()
        if not sel:
            return

        ordered = [i for i in self._sel_order if i in sel]
        if not ordered:
            ordered = sel

        parts = []
        for i in ordered:
            if 0 <= i < len(self.view_items):
                parts.append(self.view_items[i])

        if not parts:
            return

        combined = "\n".join(parts)
        self._add_history_item(combined)

        if combined in self.history:
            self._selected_item_text = combined
            self._set_preview_text(self._get_preview_display_text_for_item(combined), mark_clean=True)
            self.status_var.set(f"Combined {len(parts)} items into new entry — {now_ts()}")

    # -----------------------------
    # Find / Highlight
    # -----------------------------
    def _highlight_query_in_preview(self, query: str):
        self.preview.tag_remove("match", "1.0", tk.END)
        if not query:
            return

        text = self.preview.get("1.0", tk.END)
        if not text.strip():
            return

        pattern = re.compile(re.escape(query), re.IGNORECASE)
        m = pattern.search(text)
        if not m:
            return

        start_index = f"1.0+{m.start()}c"
        end_index = f"1.0+{m.end()}c"

        self.preview.tag_add("match", start_index, end_index)
        try:
            self.preview.tag_config("match", background="#2b78ff", foreground="white")
        except Exception:
            pass
        self.preview.see(start_index)

    def _search(self):
        q = self.search_var.get().strip()
        self.search_query = q
        self.search_matches = []
        self.search_index = 0

        if not q:
            self.status_var.set("Search cleared.")
            self.preview.tag_remove("match", "1.0", tk.END)
            return

        for item in list(self.history):
            if q.lower() in item.lower():
                self.search_matches.append(item)

        if not self.search_matches:
            self.status_var.set(f"No matches for: {q}")
            self.preview.tag_remove("match", "1.0", tk.END)
            return

        self.status_var.set(f"Found {len(self.search_matches)} matches for: {q}")
        self._jump_to_item(self.search_matches[0], highlight_query=q)

    def _jump_match(self, direction: int):
        if not self.search_matches:
            self._search()
            return
        self.search_index = (self.search_index + direction) % len(self.search_matches)
        self._jump_to_item(self.search_matches[self.search_index], highlight_query=self.search_query)

    def _jump_to_item(self, item_text: str, highlight_query: str = ""):
        if item_text not in self.view_items and self._current_filter() != "all":
            try:
                self.filter_var.set("all")
            except Exception:
                pass
            self._refresh_list()

        if item_text not in self.view_items:
            return

        idx = self.view_items.index(item_text)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.see(idx)

        self._selected_item_text = item_text
        display = self._get_preview_display_text_for_item(item_text)
        self._set_preview_text(display, mark_clean=True)

        if highlight_query:
            self._highlight_query_in_preview(highlight_query)

    # -----------------------------
    # Import / Export
    # -----------------------------
    def _export(self):
        path = filedialog.asksaveasfilename(
            title="Export History",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="copy2_export.json",
        )
        if not path:
            return

        data = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "exported_at": now_ts(),
            "history": list(self.history),
            "favorites": list(self.favorites),
            "settings": asdict(self.settings),
        }
        try:
            Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self.status_var.set(f"Exported — {path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Export failed:\n{e}")

    def _import(self):
        path = filedialog.askopenfilename(title="Import History", filetypes=[("JSON", "*.json")])
        if not path:
            return

        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            history_in = data.get("history", [])
            favorites_in = data.get("favorites", [])
            settings_in = data.get("settings", None)

            if isinstance(settings_in, dict):
                apply = messagebox.askyesno(APP_NAME, "This export includes Settings.\n\nApply imported settings too?")
                if apply:
                    new_settings = Settings.from_dict(settings_in)
                    old_theme = self.settings.theme
                    self.settings = new_settings
                    self._ensure_capacity_for_favorites(notify=False)
                    if USE_TTKB and self.settings.theme != old_theme:
                        messagebox.showinfo(APP_NAME, "Theme imported. Close and reopen the app to apply the new theme.")

            merged = list(self.history)

            if isinstance(history_in, list):
                for item in history_in:
                    if isinstance(item, str) and item.strip():
                        merged.append(item)

            favs = list(self.favorites)
            if isinstance(favorites_in, list):
                for x in favorites_in:
                    if isinstance(x, str) and x not in favs:
                        favs.append(x)

            for fav in favs:
                if fav not in merged:
                    merged.append(fav)

            seen = set()
            out = []
            for item in reversed(merged):
                if item not in seen:
                    out.append(item)
                    seen.add(item)
            out.reverse()

            self.favorites = favs
            self._ensure_capacity_for_favorites(notify=True)

            self.history = deque(out)
            ok = self._apply_capacity_preserving_favorites(reason="import")
            if not ok:
                messagebox.showwarning(
                    APP_NAME,
                    "Import completed, but favorites exceed configured capacity.\n\n"
                    "Increase Max history or remove some favorites."
                )

            self._refresh_list(select_last=True)
            self._persist()
            self.status_var.set(f"Imported — {path}")

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")

    # -----------------------------
    # Settings dialog (with Controls)
    # -----------------------------
    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        if USE_TTKB:
            Frame = tb.Frame
            Label = tb.Label
            Entry = tb.Entry
            Button = tb.Button
            Checkbutton = tb.Checkbutton
            Combobox = tb.Combobox
            Labelframe = tb.Labelframe
        else:
            from tkinter import ttk as _ttk  # type: ignore
            Frame = _ttk.Frame
            Label = _ttk.Label
            Entry = _ttk.Entry
            Button = _ttk.Button
            Checkbutton = _ttk.Checkbutton
            Combobox = _ttk.Combobox
            Labelframe = _ttk.Labelframe

        frm = Frame(dlg, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        max_var = tk.StringVar(value=str(self.settings.max_history))
        poll_var = tk.StringVar(value=str(self.settings.poll_ms))
        sess_var = tk.BooleanVar(value=self.settings.session_only)

        theme_var = tk.StringVar(value=self.settings.theme)
        themes = ["flatly", "litera", "cosmo", "sandstone", "minty", "darkly", "superhero", "cyborg"]

        Label(frm, text="Max history (5–500):").grid(row=0, column=0, sticky="w")
        Entry(frm, textvariable=max_var, width=12).grid(row=0, column=1, sticky="w", padx=(10, 0))

        Label(frm, text="Poll interval ms (100–5000):").grid(row=1, column=0, sticky="w", pady=(10, 0))
        Entry(frm, textvariable=poll_var, width=12).grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(10, 0))

        Checkbutton(frm, text="Session-only (do not save history)", variable=sess_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

        if USE_TTKB:
            Label(frm, text="Theme:").grid(row=3, column=0, sticky="w", pady=(10, 0))
            cb = Combobox(frm, textvariable=theme_var, values=themes, width=18, state="readonly")
            cb.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        else:
            Label(frm, text="(Install ttkbootstrap for modern themes)", foreground="#666").grid(
                row=3, column=0, columnspan=2, sticky="w", pady=(10, 0)
            )

        controls = Labelframe(frm, text="Controls / Hotkeys", padding=10)
        controls.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 0))

        hotkeys = (
            "Core Hotkeys:\n"
            "  Ctrl+F           Focus Search\n"
            "  Enter            Find (search)\n"
            "  Ctrl+C           Copy Preview\n"
            "  Delete           Delete selected history item\n"
            "  Ctrl+S           Save Preview Edit\n"
            "  Esc              Revert Preview Edit\n"
            "  Ctrl+E           Export\n"
            "  Ctrl+I           Import\n"
            "  Ctrl+L           Clean (remove non-favorites, keep ★)\n"
            "  Ctrl+Shift+L     Clear ALL (including favorites)\n"
            "  Ctrl+U           Check Updates\n\n"
            "History:\n"
            "  Ctrl+Click       Multi-select in click order\n"
            "  Combine          Creates a new entry (does not delete originals)\n"
        )

        txt = tk.Text(controls, height=9, wrap="word")
        txt.insert("1.0", hotkeys)
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)

        btns = Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))

        def save():
            try:
                mh = int(max_var.get().strip())
                pm = int(poll_var.get().strip())
            except Exception:
                messagebox.showerror(APP_NAME, "Please enter valid integers.")
                return

            mh = max(5, min(HARD_MAX_HISTORY, mh))
            pm = max(100, min(5000, pm))

            self.settings.max_history = mh
            self.settings.poll_ms = pm
            self.settings.session_only = bool(sess_var.get())

            if USE_TTKB:
                self.settings.theme = str(theme_var.get()).strip() or "flatly"

            self._ensure_capacity_for_favorites(notify=True)
            ok = self._apply_capacity_preserving_favorites(reason="settings_save")
            if not ok:
                messagebox.showwarning(
                    APP_NAME,
                    "Favorites exceed configured capacity.\n\nIncrease Max history or remove some favorites."
                )

            self._refresh_list(select_last=True)
            self._persist()

            if USE_TTKB and self.settings.theme != getattr(self, "_active_theme", self.settings.theme):
                messagebox.showinfo(APP_NAME, "Theme saved. Close and reopen the app to apply the new theme.")
            dlg.destroy()

        Button(btns, text="Cancel", command=dlg.destroy).grid(row=0, column=0, padx=(0, 10))
        Button(btns, text="Save", command=save).grid(row=0, column=1)

    # -----------------------------
    # Context menu
    # -----------------------------
    def _build_context_menu(self):
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copy Preview", command=self._copy_selected)
        self.menu.add_command(label="Favorite / Unfavorite", command=self._toggle_favorite_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Save Preview Edit (Ctrl+S)", command=self._save_preview_edits)
        self.menu.add_command(label="Revert Preview Edit (Esc)", command=self._revert_preview_edits)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self._delete_selected)
        self.menu.add_command(label="Combine Selected", command=self._combine_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Clean (Keep ★)", command=self._clean_history_keep_favorites)
        self.menu.add_command(label="Clear ALL (incl ★)", command=self._clear_all_history)

        def popup(event):
            try:
                idx = self.listbox.nearest(event.y)
                if idx >= 0:
                    if not (event.state & 0x0004):  # Ctrl mask
                        self.listbox.selection_clear(0, tk.END)
                        self.listbox.selection_set(idx)
                        self._prev_sel_set = set(self.listbox.curselection())
                        self._sel_order = list(self._prev_sel_set)
                    self._on_select()
            except Exception:
                pass
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    self.menu.grab_release()
                except Exception:
                    pass

        self.listbox.bind("<Button-3>", popup)


# -----------------------------
# ttkbootstrap variant
# -----------------------------
if USE_TTKB:

    class Copy2App(tb.Window, Copy2AppBase):
        def __init__(self):
            tmp_settings = Settings.from_dict(safe_json_load(Path(user_data_dir(APP_ID, VENDOR)) / "config.json", {}))
            theme = tmp_settings.theme if tmp_settings.theme else "flatly"
            super().__init__(themename=theme)
            self._active_theme = theme

            self.title(APP_NAME)
            self._init_state()
            self._apply_window_defaults()
            self._build_ui()
            self._bind_shortcuts()
            self._build_context_menu()
            self._refresh_list(select_last=True)

            self.after(250, self._poll_clipboard)

            if AUTO_CHECK_UPDATES_ON_LAUNCH:
                self.after(900, lambda: self._check_updates_async(prompt_if_new=True))

            self.protocol("WM_DELETE_WINDOW", self._on_close)

        def _apply_window_defaults(self):
            self.minsize(1100, 720)
            try:
                self.tk.call("tk", "scaling", 1.2)
            except Exception:
                pass

        def _build_ui(self):
            root = tb.Frame(self, padding=12)
            root.pack(fill=BOTH, expand=True)

            top = tb.Frame(root)
            top.pack(fill=X, pady=(0, 10))

            title_box = tb.Frame(top)
            title_box.pack(side=LEFT, padx=(0, 14))
            tb.Label(title_box, text="Copy 2.0", font=("Segoe UI", 18, "bold")).pack(anchor="w")
            self.capture_state_var = tk.StringVar(value="Capturing")
            tb.Label(title_box, textvariable=self.capture_state_var, font=("Segoe UI", 10)).pack(anchor="w")

            # Search
            search_box = tb.Frame(top)
            search_box.pack(side=LEFT, fill=X, expand=True)

            tb.Label(search_box, text="Search").pack(side=LEFT, padx=(0, 8))
            self.search_var = tk.StringVar(value="")
            self.search_entry = tb.Entry(search_box, textvariable=self.search_var)
            self.search_entry.pack(side=LEFT, fill=X, expand=True)

            btn_find = tb.Button(search_box, text="Find", command=self._search, bootstyle=PRIMARY)
            btn_find.pack(side=LEFT, padx=(10, 0))
            btn_prev = tb.Button(search_box, text="Prev", command=lambda: self._jump_match(-1), bootstyle=SECONDARY)
            btn_prev.pack(side=LEFT, padx=(8, 0))
            btn_next = tb.Button(search_box, text="Next", command=lambda: self._jump_match(1), bootstyle=SECONDARY)
            btn_next.pack(side=LEFT, padx=(8, 0))

            self._tip(btn_find, "Find (Enter)")
            self._tip(btn_prev, "Previous match")
            self._tip(btn_next, "Next match")

            tb.Frame(top, width=18).pack(side=LEFT)

            action_box = tb.Frame(top)
            action_box.pack(side=RIGHT)

            self.pause_var = tk.BooleanVar(value=self.paused)
            chk_pause = tb.Checkbutton(
                action_box,
                text="Pause",
                variable=self.pause_var,
                command=self._toggle_pause,
                bootstyle="round-toggle",
            )
            chk_pause.pack(side=LEFT, padx=(10, 18))
            self._tip(chk_pause, "Pause clipboard capture")

            btn_updates = tb.Button(
                action_box,
                text="Check Updates",
                command=lambda: self._check_updates_async(prompt_if_new=False),
                bootstyle=OUTLINE
            )
            btn_updates.pack(side=LEFT, padx=(0, 8))
            self._tip(btn_updates, "Check for updates now (Ctrl+U)")

            btn_export = tb.Button(action_box, text="Export", command=self._export, bootstyle=OUTLINE)
            btn_export.pack(side=LEFT, padx=(0, 8))
            btn_import = tb.Button(action_box, text="Import", command=self._import, bootstyle=OUTLINE)
            btn_import.pack(side=LEFT, padx=(0, 8))
            btn_settings = tb.Button(action_box, text="Settings", command=self._open_settings, bootstyle=SECONDARY)
            btn_settings.pack(side=LEFT)

            self._tip(btn_export, "Export (Ctrl+E)")
            self._tip(btn_import, "Import (Ctrl+I)")
            self._tip(btn_settings, "Settings")

            tb.Separator(root).pack(fill=X, pady=(0, 10))

            body = tb.Frame(root)
            body.pack(fill=BOTH, expand=True)
            body.columnconfigure(0, weight=1)
            body.columnconfigure(1, weight=2)
            body.rowconfigure(0, weight=1)

            # Left pane
            left = tb.Labelframe(body, text="History", padding=10)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            left.rowconfigure(2, weight=1)
            left.columnconfigure(0, weight=1)

            filters = tb.Frame(left)
            filters.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            filters.columnconfigure(3, weight=1)

            self.filter_var = tk.StringVar(value="all")
            tb.Radiobutton(filters, text="All", value="all", variable=self.filter_var,
                           command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=0, padx=(0, 8))
            tb.Radiobutton(filters, text="Favorites", value="fav", variable=self.filter_var,
                           command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=1, padx=(0, 8))

            btn_clean = tb.Button(filters, text="Clean", command=self._clear_history, bootstyle=DANGER)
            btn_clean.grid(row=0, column=4, sticky="e")
            self._tip(btn_clean, "Clean (Ctrl+L) — keeps ★")

            self.listbox = tk.Listbox(left, activestyle="dotbox", exportselection=False, selectmode=tk.EXTENDED)
            self.listbox.grid(row=2, column=0, sticky="nsew")
            self.listbox.bind("<<ListboxSelect>>", self._on_listbox_select_event)
            self.listbox.bind("<Double-Button-1>", lambda e: self._copy_selected())

            sb = tb.Scrollbar(left, orient="vertical", command=self.listbox.yview)
            sb.grid(row=2, column=1, sticky="ns")
            self.listbox.configure(yscrollcommand=sb.set)

            left_actions = tb.Frame(left)
            left_actions.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
            for i in range(4):
                left_actions.columnconfigure(i, weight=1)

            tb.Button(left_actions, text="Copy", command=self._copy_selected, bootstyle=SUCCESS)\
                .grid(row=0, column=0, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Delete", command=self._delete_selected, bootstyle=WARNING)\
                .grid(row=0, column=1, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Fav / Unfav", command=self._toggle_favorite_selected, bootstyle=INFO)\
                .grid(row=0, column=2, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Combine", command=self._combine_selected, bootstyle=PRIMARY)\
                .grid(row=0, column=3, sticky="ew")

            # Right pane
            right = tb.Labelframe(body, text="Preview (Editable)", padding=10)
            right.grid(row=0, column=1, sticky="nsew")
            right.rowconfigure(1, weight=1)
            right.columnconfigure(0, weight=1)

            preview_actions = tb.Frame(right)
            preview_actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            preview_actions.columnconfigure(0, weight=1)

            self.reverse_var = tk.BooleanVar(value=self.settings.reverse_lines_copy)
            tb.Checkbutton(
                preview_actions,
                text="Reverse-lines copy",
                variable=self.reverse_var,
                command=self._toggle_reverse_lines,
                bootstyle="round-toggle",
            ).pack(side=LEFT)

            self.preview_dirty_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.preview_dirty_var).pack(side=LEFT, padx=(10, 0))

            self.revert_btn = tb.Button(preview_actions, text="Revert", command=self._revert_preview_edits, bootstyle=WARNING)
            self.revert_btn.pack(side=RIGHT, padx=(0, 8))
            self.save_btn = tb.Button(preview_actions, text="Save Edit", command=self._save_preview_edits, bootstyle=SUCCESS)
            self.save_btn.pack(side=RIGHT, padx=(0, 8))
            tb.Button(preview_actions, text="Copy Preview", command=self._copy_selected, bootstyle=SUCCESS).pack(side=RIGHT)

            self.preview = tk.Text(right, wrap="word", undo=True)
            self.preview.grid(row=1, column=0, sticky="nsew")
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            sb2 = tb.Scrollbar(right, orient="vertical", command=self.preview.yview)
            sb2.grid(row=1, column=1, sticky="ns")
            self.preview.configure(yscrollcommand=sb2.set)

            bottom = tb.Frame(root)
            bottom.pack(fill=X, pady=(10, 0))
            self.status_var = tk.StringVar(value="")
            tb.Label(bottom, textvariable=self.status_var).pack(side=LEFT)

            self._update_preview_dirty_ui()

        def _bind_shortcuts(self):
            self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
            self.bind("<Return>", lambda e: (self._search(), "break"))
            self.bind("<Control-c>", lambda e: (self._copy_selected(), "break"))
            self.bind("<Delete>", lambda e: (self._delete_selected(), "break"))
            self.bind("<Control-e>", lambda e: (self._export(), "break"))
            self.bind("<Control-i>", lambda e: (self._import(), "break"))
            self.bind("<Control-l>", lambda e: (self._clear_history(), "break"))
            self.bind("<Control-Shift-L>", lambda e: (self._clear_all_history(), "break"))
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))
            self.bind("<Control-u>", lambda e: (self._check_updates_async(prompt_if_new=False), "break"))

        def _on_close(self):
            try:
                if self._poll_job is not None:
                    self.after_cancel(self._poll_job)
            except Exception:
                pass

            if self._preview_dirty:
                resp = messagebox.askyesnocancel(
                    APP_NAME,
                    "You have unsaved edits in Preview.\n\n"
                    "Yes = Save edits\nNo = Discard edits\nCancel = Keep app open"
                )
                if resp is None:
                    return
                if resp is True:
                    self._save_preview_edits()

            self._persist()
            self.destroy()

else:
    # ttk fallback (minimal)
    class Copy2App(tk.Tk, Copy2AppBase):
        def __init__(self):
            super().__init__()
            self.title(APP_NAME)
            self._init_state()
            self.minsize(1100, 720)

            self._build_ui()
            self._bind_shortcuts()
            self._build_context_menu()
            self._refresh_list(select_last=True)

            self.after(250, self._poll_clipboard)
            if AUTO_CHECK_UPDATES_ON_LAUNCH:
                self.after(900, lambda: self._check_updates_async(prompt_if_new=True))
            self.protocol("WM_DELETE_WINDOW", self._on_close)

        def _build_ui(self):
            from tkinter import ttk
            root = ttk.Frame(self, padding=12)
            root.pack(fill=tk.BOTH, expand=True)
            ttk.Label(root, text="Please install ttkbootstrap for the modern UI.").pack()

            self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED)
            self.listbox.pack(fill=tk.BOTH, expand=True)
            self.listbox.bind("<<ListboxSelect>>", self._on_listbox_select_event)

            self.preview = tk.Text(root, wrap="word", undo=True)
            self.preview.pack(fill=tk.BOTH, expand=True)
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            self.search_var = tk.StringVar(value="")
            self.filter_var = tk.StringVar(value="all")
            self.pause_var = tk.BooleanVar(value=False)
            self.reverse_var = tk.BooleanVar(value=self.settings.reverse_lines_copy)
            self.status_var = tk.StringVar(value=f"v{APP_VERSION}")

            self.preview_dirty_var = tk.StringVar(value="")
            self.save_btn = None
            self.revert_btn = None

        def _bind_shortcuts(self):
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))
            self.bind("<Control-l>", lambda e: (self._clear_history(), "break"))
            self.bind("<Control-Shift-L>", lambda e: (self._clear_all_history(), "break"))
            self.bind("<Control-u>", lambda e: (self._check_updates_async(prompt_if_new=False), "break"))

        def _on_close(self):
            try:
                if self._poll_job is not None:
                    self.after_cancel(self._poll_job)
            except Exception:
                pass
            self._persist()
            self.destroy()


def main():
    app = Copy2App() if not USE_TTKB else Copy2App()
    app.mainloop()


if __name__ == "__main__":
    main()
