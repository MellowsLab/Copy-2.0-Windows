"""
Copy 2.0 (Windows)

Key fixes added:
- Favorites are NEVER removed by "Clean" or by history-cap pruning.
- When max history is reached, you are notified to increase it.
- Hard cap protection (HARD_MAX_HISTORY): if you hit it and cannot prune (because favorites), you are warned.
- Settings now includes a Controls/Hotkeys section.
- Robust GitHub update checker (startup + button):
  - Uses GitHub Releases API
  - Falls back to parsing release body for .zip links (including user-attachments)
- Robust auto-update installer:
  - Downloads ZIP/EXE
  - Stages Copy2.exe
  - Writes a .bat that:
      - waits for current PID to exit
      - backs up current EXE
      - copies new EXE into place
      - clears PyInstaller/Python env vars
      - runs self-test with retries
      - rolls back automatically on failure
  - Logs: update_check.log, update_install.log, update_bat.log

Dependencies:
  pyperclip
  platformdirs
Optional (recommended UI):
  ttkbootstrap

Build (onefile recommended for auto-update):
  python -m PyInstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py

Note:
- Auto-update is ONLY supported for frozen onefile builds.
- If you are running an onedir build (folder next to exe), updates will open the release page instead.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import zipfile
import tempfile
import threading
import subprocess
import shutil
import webbrowser
import base64
import hashlib
import secrets
import ctypes

# --- Windows taskbar identity (must be set before creating any Tk window) ---
if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MellowLabs.Copy2")
    except Exception:
        pass

# Optional: imaging (clipboard images + screenshots)
try:
    from PIL import Image, ImageTk, ImageGrab  # type: ignore
except Exception:
    Image = ImageTk = ImageGrab = None

# Optional: encryption for exports
try:
    from cryptography.fernet import Fernet  # type: ignore
except Exception:
    Fernet = None

# Optional: key derivation for encryption-at-rest and encrypted exports
try:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC  # type: ignore
    from cryptography.hazmat.primitives import hashes  # type: ignore
except Exception:
    PBKDF2HMAC = None
    hashes = None

# Optional: Windows start-on-boot
try:
    import winreg  # type: ignore
except Exception:
    winreg = None
from collections import deque
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk

# Optional: global hotkeys + typing (Windows)
try:
    import keyboard as _kbd  # type: ignore
except Exception:
    _kbd = None

# Optional: clipboard helper (recommended).
# If pyperclip is unavailable, we fall back to Tk clipboard APIs.
try:
    import pyperclip  # type: ignore
except Exception:
    pyperclip = None

from platformdirs import user_data_dir

# -----------------------------
# Optional modern theme layer
# -----------------------------
USE_TTKB = True
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
except Exception:
    USE_TTKB = False
    # ttk is imported unconditionally above; keep fallback constants for non-ttkbootstrap.

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

# -----------------------------
# App constants
# -----------------------------
APP_NAME = "Copy 2.0"
APP_ID = "copy2"
VENDOR = "MellowsLab"

# App version (display + update compare)
APP_VERSION = "1.0.8"  # <-- keep this in sync with your build/tag

# History indicator icons (you can change these to any emoji you like)
PIN_ICON = "📌"
FAV_ICON = "⭐"
TAG_ICON = "🏷️"
IMAGE_ICON = "🖼️"

DEFAULT_MAX_HISTORY = 50
DEFAULT_POLL_MS = 400

# Hard cap: beyond this, the UI will not allow increasing max_history
# and the app will warn that you have used all allocated memory.
HARD_MAX_HISTORY = 500

# GitHub update config
GITHUB_OWNER = "MellowsLab"
GITHUB_REPO = "Copy-2.0-Windows"
GITHUB_LATEST_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

# -----------------------------
# Utilities
# -----------------------------
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


def _norm_ver(v: str) -> tuple[int, int, int]:
    """
    Normalize versions like:
      "v1.2.3" -> (1,2,3)
      "1.2"   -> (1,2,0)
      "1"     -> (1,0,0)
    """
    v = (v or "").strip()
    v = v[1:] if v.lower().startswith("v") else v
    parts = re.findall(r"\d+", v)
    nums = [int(x) for x in parts[:3]] + [0, 0, 0]
    return (nums[0], nums[1], nums[2])


def is_newer_version(latest: str, installed: str) -> bool:
    return _norm_ver(latest) > _norm_ver(installed)


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def is_onedir_frozen(exe_path: Path) -> bool:
    """
    Best-effort detection: if common runtime files exist next to exe,
    it's likely an onedir build. Onefile typically does NOT have these.
    """
    parent = exe_path.parent
    markers = [
        "python312.dll",
        "python311.dll",
        "python310.dll",
        "base_library.zip",
        "libcrypto-3.dll",
        "libssl-3.dll",
    ]
    return any((parent / m).exists() for m in markers)


# -----------------------------
# Simple tooltip (hoverover hotkeys)
# -----------------------------
class ToolTip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 10
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self.tip = tk.Toplevel(self.widget)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{x}+{y}")
            lbl = tk.Label(
                self.tip,
                text=self.text,
                justify="left",
                background="#1f1f1f",
                foreground="white",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=4,
                font=("Segoe UI", 9),
            )
            lbl.pack()
        except Exception:
            self.tip = None

    def _hide(self, _event=None):
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


# -----------------------------
# Settings
# -----------------------------
@dataclass
class Settings:
    max_history: int = DEFAULT_MAX_HISTORY
    poll_ms: int = DEFAULT_POLL_MS
    session_only: bool = False
    reverse_lines_copy: bool = False
    theme: str = "flatly"
    check_updates_on_launch: bool = True

    # UI
    pane_sash: int = 420  # initial left pane width (px)

    # Global hotkeys (optional, requires 'keyboard' package)
    enable_global_hotkeys: bool = False
    hotkey_quick_paste: str = "ctrl+alt+v"
    hotkey_paste_last: str = "ctrl+alt+shift+v"
    hotkey_toggle_pause: str = "ctrl+alt+p"

    # File-based sync (point to a cloud-synced folder like OneDrive/Dropbox)
    sync_enabled: bool = False
    sync_folder: str = ""
    sync_interval_sec: int = 10

    # Advanced features (ALL disabled by default)
    advanced_features: bool = True
    adv_app_lock: bool = False
    # Auto-lock UI after inactivity (minutes). 0 = never.
    lock_timeout_minutes: int = 0
    adv_start_on_boot: bool = True
    adv_encrypt_exports: bool = False
    adv_encrypt_all_data: bool = False
    adv_images: bool = False
    adv_screenshots: bool = False
    adv_snippets: bool = False
    adv_tmplt_trigger: bool = False
    tmplt_trigger_word: str = "tmplt"

    @staticmethod
    def from_dict(d: dict) -> "Settings":
        s = Settings()
        s.max_history = int(d.get("max_history", DEFAULT_MAX_HISTORY))
        s.poll_ms = int(d.get("poll_ms", DEFAULT_POLL_MS))
        s.session_only = bool(d.get("session_only", False))
        s.reverse_lines_copy = bool(d.get("reverse_lines_copy", False))
        s.theme = str(d.get("theme", "flatly"))
        s.check_updates_on_launch = bool(d.get("check_updates_on_launch", True))

        s.pane_sash = int(d.get("pane_sash", 420))

        s.enable_global_hotkeys = bool(d.get("enable_global_hotkeys", False))
        s.hotkey_quick_paste = str(d.get("hotkey_quick_paste", "ctrl+alt+v"))
        s.hotkey_paste_last = str(d.get("hotkey_paste_last", "ctrl+alt+shift+v"))
        s.hotkey_toggle_pause = str(d.get("hotkey_toggle_pause", "ctrl+alt+p"))

        s.sync_enabled = bool(d.get("sync_enabled", False))
        s.sync_folder = str(d.get("sync_folder", "")).strip()
        s.sync_interval_sec = int(d.get("sync_interval_sec", 10))

        # Defaults: Advanced is enabled by default, and Start-on-boot is enabled by default.
        # Other advanced switches remain opt-in.
        s.advanced_features = bool(d.get("advanced_features", True))
        s.adv_app_lock = bool(d.get("adv_app_lock", False))
        s.lock_timeout_minutes = int(d.get("lock_timeout_minutes", 0))
        s.adv_start_on_boot = bool(d.get("adv_start_on_boot", True))
        s.adv_encrypt_exports = bool(d.get("adv_encrypt_exports", False))
        s.adv_encrypt_all_data = bool(d.get("adv_encrypt_all_data", False))
        s.adv_images = bool(d.get("adv_images", False))
        s.adv_screenshots = bool(d.get("adv_screenshots", False))
        s.adv_snippets = bool(d.get("adv_snippets", False))
        s.adv_tmplt_trigger = bool(d.get("adv_tmplt_trigger", False))
        s.tmplt_trigger_word = str(d.get("tmplt_trigger_word", "tmplt"))

        s.max_history = max(5, min(HARD_MAX_HISTORY, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        s.pane_sash = max(220, min(900, s.pane_sash))
        s.lock_timeout_minutes = max(0, min(24*60, int(getattr(s, "lock_timeout_minutes", 0))))
        s.sync_interval_sec = max(3, min(300, s.sync_interval_sec))
        if not s.tmplt_trigger_word.strip():
            s.tmplt_trigger_word = "tmplt"
        return s




def _normalize_text_block(obj) -> str:
    """Return a safe string for UI insertion.

    Accepts:
    - str
    - list/tuple of str (implicit concatenation blocks)
    - anything else (str(obj))
    """
    if obj is None:
        return ""
    if isinstance(obj, str):
        return obj
    if isinstance(obj, (list, tuple)):
        try:
            return "".join([x if isinstance(x, str) else str(x) for x in obj])
        except Exception:
            return ""
    try:
        return str(obj)
    except Exception:
        return ""
# -----------------------------
# Help text (shown in Settings → Help)
# -----------------------------
COPY2_HELP_TEXT = '''
OVERVIEW
- Copy2 watches your clipboard and stores recent items into the History list.
- You can search, preview, copy back to clipboard, and manage items (favorite / pin / tag / expiry).

HISTORY (MAIN LIST)
- The History list shows captured clipboard items (usually newest first).
- Click an item to load it into Preview.
- Double-click copies the selected item back to clipboard.

RIGHT-CLICK IN HISTORY (CONTEXT MENU)
- Copy: Copies the selected item back to clipboard.
- Favorite / Unfavorite: Toggles “favorite” status for the selected item.
- Pin / Unpin: Toggles “pinned” status for the selected item (pinned items should stay protected).
- Tags…: Add/remove tags on the selected item.
- Expiry…: Set/clear an expiry time for the selected item.
- Delete: Removes the selected item from history.
- Combine Selected: Combines multiple selected items into one entry.

SEARCH
- The Search box filters the History list as you type.
- Clearing search returns to the full list.

PREVIEW (DETAIL VIEW)
- Preview shows the full content of the selected history item.
- Some builds allow editing in Preview and then “Save Preview Edit” / “Revert Preview Edit”.

FAVORITES
- Favorites mark important entries you want to keep.
- Favorites should not be removed by “Clean”.
- Typical indicator: a star icon or star marker.

PINS
- Pins are stronger protection than favorites: “don’t let this disappear”.
- Pinned items should not be removed by “Clean”.
- Typical indicator: a pin icon.

TAGS
- Tags help categorize items (e.g., Work, Code, Passwords, Clips).
- Use the Tag dropdown to filter by tag (no separate “show all tags” button needed).
- Tags should be visible and easy to manage.

EXPIRY
- Expiry automatically removes selected items after a set time in minutes.

CLEAN
- “Clean” removes unprotected clutter items.
- Intended rule: items that are favorited / pinned / tagged should be protected from cleaning.

QUICK PASTE / TYPE
- Quick Paste is a fast picker window: search history and either:
  - Copy: copy selected item to clipboard
  - Paste: paste into the active app, way to use this is keep the quick paste window open, click where you want to paste, double click which text you want to paste.
  - Type: simulate typing the item into the active app, this should get around fields of entry that dont support pasting into.
- Global hotkeys (if enabled) open Quick Paste instantly.

SYNC (CLOUD SYNC)
- Sync (if enabled) keeps your history/favorites/pins/tags consistent across machines, this is done with onedrive or any other form of file cloud syncing.

SETTINGS
- Settings generally control capture behavior, storage limits, theme/UI options, hotkeys, and sync.
'''


# -----------------------------
# GitHub update helpers
# -----------------------------
def _http_get_json(url: str, timeout: int = 10) -> dict | None:
    try:
        import urllib.request
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": f"{APP_NAME}/{APP_VERSION} ({sys.platform})",
                "Accept": "application/vnd.github+json",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        return json.loads(data.decode("utf-8", errors="ignore"))
    except Exception:
        return None


def _http_download(url: str, out_path: Path, timeout: int = 30) -> tuple[int, int | None]:
    """
    Returns: (downloaded_bytes, expected_bytes or None)
    """
    import urllib.request

    out_path.parent.mkdir(parents=True, exist_ok=True)

    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"{APP_NAME}/{APP_VERSION} ({sys.platform})",
            "Accept": "*/*",
        },
    )

    expected = None
    downloaded = 0
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        try:
            cl = resp.headers.get("Content-Length")
            if cl:
                expected = int(cl)
        except Exception:
            expected = None

        with open(out_path, "wb") as f:
            while True:
                chunk = resp.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)

    return downloaded, expected


def _parse_zip_url_from_body(body: str) -> str | None:
    """
    Fallback when release assets are empty (e.g., "user-attachments" links).
    Extract the first URL ending with .zip (or .exe) from the release body.
    """
    if not body:
        return None

    # Match markdown links and raw URLs
    urls = re.findall(r"(https?://[^\s)]+)", body, flags=re.IGNORECASE)
    # prefer .zip
    for u in urls:
        if u.lower().endswith(".zip"):
            return u
    # fallback .exe
    for u in urls:
        if u.lower().endswith(".exe"):
            return u
    return None


def get_latest_release_info() -> dict | None:
    """
    Returns dict with:
      version: "v1.2.3"
      html_url: release page
      asset_url: direct download url (.zip preferred)
      asset_name: filename
    """
    j = _http_get_json(GITHUB_LATEST_API, timeout=10)
    if not isinstance(j, dict):
        return None

    tag = str(j.get("tag_name") or "").strip()
    html_url = str(j.get("html_url") or GITHUB_RELEASES_PAGE).strip()

    asset_url = ""
    asset_name = ""

    assets = j.get("assets", [])
    if isinstance(assets, list):
        # Prefer .zip assets first
        zip_assets = [a for a in assets if str(a.get("name", "")).lower().endswith(".zip")]
        exe_assets = [a for a in assets if str(a.get("name", "")).lower().endswith(".exe")]

        chosen = None
        if zip_assets:
            # prefer a common name if present
            preferred = None
            for a in zip_assets:
                nm = str(a.get("name", "")).lower()
                if nm in ("copy2.zip", "copy-2.0.zip", "copy2win.zip"):
                    preferred = a
                    break
            chosen = preferred or zip_assets[0]
        elif exe_assets:
            chosen = exe_assets[0]

        if chosen:
            asset_url = str(chosen.get("browser_download_url") or "").strip()
            asset_name = str(chosen.get("name") or "").strip()

    # Fallback: parse release body for a zip/exe link
    if not asset_url:
        body = str(j.get("body") or "")
        u = _parse_zip_url_from_body(body)
        if u:
            asset_url = u
            asset_name = u.split("/")[-1]

    return {
        "version": tag,
        "html_url": html_url,
        "asset_url": asset_url,
        "asset_name": asset_name,
    }


# -----------------------------
# Main app logic
# -----------------------------
class Copy2AppBase:
    """Shared logic for both ttkbootstrap and ttk fallback variants."""
    def _open_help(self, parent=None):
        """Open Help inside the Settings dialog."""
        try:
            self._open_settings(initial_tab='Help')
        except TypeError:
            # Back-compat if _open_settings signature differs
            self._open_settings()

    # -----------------------------
    # App icon (window/taskbar) helpers
    # -----------------------------
    def _get_app_icon_photo(self):
        """Return a Tk PhotoImage for the app icon (best-effort).

        Notes:
        - Prefer PNG (native tk.PhotoImage support).
        - If only an ICO is available, fall back to Pillow (ImageTk).
        """
        if getattr(self, '_app_icon_photo', None) is not None:
            return self._app_icon_photo

        img = None
        try:
            candidates = []

            def add_candidates(base: Path):
                # Common filenames (keep existing ones for backward compatibility)
                candidates.extend([
                base / "assets" / "Mellowlabs.png",
                base / "assets" / "Mellowlabs.ico",
                ])

            # PyInstaller bundle root
            try:
                meipass = getattr(sys, '_MEIPASS', None)
                if meipass:
                    add_candidates(Path(meipass))
            except Exception:
                pass

            #  с exe (onedir) / installed location
            try:
                exe_dir = Path(sys.executable).resolve().parent
                add_candidates(exe_dir)
            except Exception:
                pass

            # dev run (next to this .py)
            try:
                here = Path(__file__).resolve().parent
                add_candidates(here)
            except Exception:
                pass

            # 1) Try PNG first (fast, native)
            for p in candidates:
                try:
                    if p.exists() and p.suffix.lower() == '.png':
                        img = tk.PhotoImage(file=str(p))
                        break
                except Exception:
                    continue
        
            # 2) Fall back to ICO via Pillow
            if img is None:
                for p in candidates:
                    try:
                        if p.exists() and p.suffix.lower() == '.ico' and Image is not None and ImageTk is not None:
                            pil = Image.open(str(p))
                            # Use the largest available frame/size when ICO contains multiple sizes.
                            try:
                                if getattr(pil, 'n_frames', 1) > 1:
                                    pil.seek(pil.n_frames - 1)
                            except Exception:
                                pass
                            img = ImageTk.PhotoImage(pil)
                            break
                    except Exception:
                        continue
        except Exception:
            img = None

        self._app_icon_photo = img
        return img
    def _apply_app_icon_to_window(self):
    #Apply icon to the main window/taskbar (best-effort).
    # Titlebar/top-left (ICO)
        try:
            ico_candidates = []

            def add(base: Path):
                ico_candidates.extend([
                    base / "assets" / "Mellowlabs.ico",
                    base / "Mellowlabs.ico",
                    base / "assets" / "copy2.ico",
                    base / "copy2.ico",
                ])

            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                add(Path(meipass))

            add(Path(sys.executable).resolve().parent)
            add(Path(__file__).resolve().parent)

            for p in ico_candidates:
                if p.exists():
                    self.iconbitmap(str(p))
                    break
        except Exception:
            pass

    # Taskbar/window icon (PNG preferred)
    try:
        png_candidates = []

        def addp(base: Path):
            png_candidates.extend([
                base / "assets" / "Mellowlabs.png",
                base / "Mellowlabs.png",
                base / "assets" / "copy2.png",
                base / "copy2.png",
                base / "assets" / "Image 4.png",
                base / "Image 4.png",
            ])

        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            addp(Path(meipass))

        addp(Path(sys.executable).resolve().parent)
        addp(Path(__file__).resolve().parent)

        for p in png_candidates:
            if p.exists() and p.suffix.lower() == ".png":
                img = tk.PhotoImage(file=str(p))
                self.iconphoto(True, img)
                self._app_icon_photo = img  # keep reference alive
                break
    except Exception:
        pass

    

    def _get_lock_logo_photo(self, max_size: int = 96):
        """Return a smaller PhotoImage for the lock screen logo.

        Uses the app icon image, but downsizes it so it cannot overlap the unlock UI.
        Best-effort; falls back to the raw icon if resizing is not possible.
        """
        if getattr(self, '_lock_logo_photo', None) is not None:
            return self._lock_logo_photo

        try:
            ico = self._get_app_icon_photo()
        except Exception:
            ico = None

        if ico is None:
            self._lock_logo_photo = None
            return None

        # If we have Pillow available, do a true resize for consistent results.
        try:
            from PIL import Image, ImageTk  # type: ignore
            # Try to locate the PNG used for the icon so we can resize it properly.
            # If not found, fall back to subsample.
            src_path = None
            try:
                candidates = []
                def add_candidates(base: Path):
                    candidates.extend([base / "assets" / "Mellowlabs.png"])
                try:
                    meipass = getattr(sys, '_MEIPASS', None)
                    if meipass:
                        add_candidates(Path(meipass))
                except Exception:
                    pass
                try:
                    exe_dir = Path(sys.executable).resolve().parent
                    add_candidates(exe_dir)
                except Exception:
                    pass
                try:
                    here = Path(__file__).resolve().parent
                    add_candidates(here)
                except Exception:
                    pass
                for p in candidates:
                    try:
                        if p.exists():
                            src_path = p
                            break
                    except Exception:
                        continue
            except Exception:
                src_path = None

            if src_path is not None:
                im = Image.open(str(src_path)).convert("RGBA")
                w, h = im.size
                scale = max(w / float(max_size), h / float(max_size), 1.0)
                nw, nh = max(1, int(w / scale)), max(1, int(h / scale))
                im = im.resize((nw, nh), Image.LANCZOS)
                self._lock_logo_photo = ImageTk.PhotoImage(im)
                return self._lock_logo_photo
        except Exception:
            pass

        # Fallback: tk.PhotoImage subsample (rough but safe)
        try:
            w = int(ico.width())
            h = int(ico.height())
            factor = int(math.ceil(max(w, h) / float(max_size)))
            if factor < 1:
                factor = 1
            if factor > 1:
                small = ico.subsample(factor, factor)
            else:
                small = ico
            self._lock_logo_photo = small
            return self._lock_logo_photo
        except Exception:
            self._lock_logo_photo = ico
            return ico

    def _init_state(self):
        self.data_dir = Path(user_data_dir(APP_ID, VENDOR))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"
        self.pins_path = self.data_dir / "pins.json"
        self.tags_path = self.data_dir / "tags.json"
        self.tag_colors_path = self.data_dir / "tag_colors.json"
        self.expiry_path = self.data_dir / "expiry.json"
        # Rich text/format preservation store (Windows clipboard HTML/RTF formats)
        self.formats_path = self.data_dir / "formats.json"
        # Advanced feature storage (optional)
        self.security_path = self.data_dir / "security.json"
        self.images_dir = self.data_dir / "images"
        self.images_meta_path = self.data_dir / "images.json"
        self.snippets_path = self.data_dir / "snippets.json"


        # Sync folder state
        self._sync_mtimes = {}
        self._sync_job = None


        self.log_update_check = self.data_dir / "update_check.log"
        self.log_update_install = self.data_dir / "update_install.log"

        self.settings = Settings.from_dict(safe_json_load(self.settings_path, {}))

        # Rich formats map: key = the stored text value, value = dict with optional html_b64/rtf_b64
        self.clip_formats = {}

        # Advanced runtime state (optional features are OFF by default)
        sec = safe_json_load(self.security_path, {})
        self.security = sec if isinstance(sec, dict) else {}
        self._unlocked = True  # set to False on startup if App Lock is enabled and a PIN exists
        self._pin_session_verified = False

        # Inactivity auto-lock (UI-only)
        self._idle_after_id = None
        self._idle_last_activity_ts = 0.0

        # If Encrypt-All is enabled and a PIN is required, defer loading encrypted stores until unlock.
        self._stores_loaded = True
        self._captured_while_locked = deque()
        try:
            if self._enc_all_enabled() and self._startup_requires_unlock():
                self._stores_loaded = False
        except Exception:
            self._stores_loaded = True


        if self._stores_loaded:
                # Images (optional) — use store loader so Encrypt-All files don't break parsing
                meta = self._store_load_json(self.images_meta_path, [])
                self.images = meta if isinstance(meta, list) else []
                self._last_image_sig = ''

                # Snippets/Templates (optional) — use store loader so Encrypt-All files don't break parsing
                sn = self._store_load_json(self.snippets_path, {'templates': []})
                if isinstance(sn, dict) and isinstance(sn.get('templates'), list):
                    self.snippets = sn.get('templates')
                elif isinstance(sn, list):
                    self.snippets = sn
                else:
                    self.snippets = []

                # Template trigger hook state
                self._tmplt_hook = None
                self._tmplt_buffer = ''
                self._tmplt_last_ts = 0.0

                # Normalize stored favorites to unique list
                favs = self._store_load_json(self.favs_path, [])
                if isinstance(favs, list):
                    favs = [x for x in favs if isinstance(x, str)]
                else:
                    favs = []
                # preserve order, unique
                seen = set()
                self.favorites = []
                for x in favs:
                    if x not in seen:
                        self.favorites.append(x)
                        seen.add(x)

                # Pins
                pins = self._store_load_json(self.pins_path, [])
                if isinstance(pins, list):
                    pins = [x for x in pins if isinstance(x, str)]
                else:
                    pins = []
                self.pins = []
                seenp = set()
                for x in pins:
                    if x not in seenp:
                        self.pins.append(x)
                        seenp.add(x)

                # Tags: {clip_text: [tag, ...]}
                tags = self._store_load_json(self.tags_path, {})
                if not isinstance(tags, dict):
                    tags = {}
                norm_tags = {}
                for k,v in tags.items():
                    if not isinstance(k, str):
                        continue
                    if isinstance(v, list):
                        vals = [str(t).strip() for t in v if str(t).strip()]
                    else:
                        vals = []
                    # unique preserve order
                    seen_t=set()
                    out=[]
                    for t in vals:
                        if t not in seen_t:
                            out.append(t)
                            seen_t.add(t)
                    if out:
                        norm_tags[k]=out
                self.tags = norm_tags

                # Tag colors: {tag_name: '#RRGGBB'}
                tc = self._store_load_json(self.tag_colors_path, {})
                if not isinstance(tc, dict):
                    tc = {}
                norm_tc = {}
                for k, v in tc.items():
                    if not isinstance(k, str) or not isinstance(v, str):
                        continue
                    kk = k.strip()
                    vv = v.strip()
                    if kk and vv:
                        norm_tc[kk] = vv
                self.tag_colors = norm_tc

                # Expiry: {clip_text: unix_ts}
                exp = self._store_load_json(self.expiry_path, {})
                if not isinstance(exp, dict):
                    exp = {}
                norm_exp = {}
                for k,v in exp.items():
                    if not isinstance(k, str):
                        continue
                    try:
                        ts = float(v)
                        norm_exp[k] = ts
                    except Exception:
                        pass
                self.expiry = norm_exp

                self._last_expiry_purge = 0.0

                hist = self._store_load_json(self.history_path, [])
                if isinstance(hist, list):
                    hist = [x for x in hist if isinstance(x, str)]
                else:
                    hist = []

                self.history = deque(hist, maxlen=self.settings.max_history)
        else:
            # Deferred stores (locked + encrypt-all). Initialize empty runtime state and load on unlock.
            self.images = []
            self._last_image_sig = ''
            self.snippets = []
            self.favorites = []
            self.pins = []
            self.tags = {}
            self.tag_colors = {}
            self.expiry = {}
            self._last_expiry_purge = 0.0
            self.history = deque([], maxlen=self.settings.max_history)

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
        self._selected_item_text: str | None = None  # original selected item
        self._reverse_item_text: str | None = None  # which item has reverse-lines applied
        self._preview_dirty = False

        # One-shot notifications
        self._warned_limit_reached = False
        self._warned_fav_block = False

        # Hotkeys map (shown in settings)
        self.HOTKEYS = {
            "Focus Search": "Ctrl+F",
            "Find": "Enter",
            "Copy Preview": "Ctrl+C",
            "Delete Selected": "Del",
            "Export": "Ctrl+E",
            "Import": "Ctrl+I",
            "Clean (keep favorites)": "Ctrl+L",
            "Save Preview Edit": "Ctrl+S",
            "Revert Preview Edit": "Esc",
            "Quick Paste (global)": "Ctrl+Alt+V",
            "Paste Last (global)": "Ctrl+Alt+Shift+V",
            "Toggle Pause (global)": "Ctrl+Alt+P",
            "Check Updates": "(Button)",
        }

        # Advanced feature bootstrapping (does not enable anything by default)
        try:
            self._advanced_bootstrap()
        except Exception:
            pass

        # Ensure favorites remain present in history (optional)
        self._ensure_favorites_present()


    # -----------------------------
    # Advanced features (disabled by default)
    # -----------------------------
    def _adv(self, key: str) -> bool:
        """Return True only if Advanced Features are enabled and the specific flag is enabled."""
        try:
            if not bool(getattr(self.settings, 'advanced_features', False)):
                return False
            return bool(getattr(self.settings, key, False))
        except Exception:
            return False

    def _advanced_bootstrap(self):
        """Create storage and register optional hooks. Never turns features on automatically."""
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        # Ensure images folder exists when the feature is enabled
        if self._adv('adv_images') or self._adv('adv_screenshots'):
            try:
                self.images_dir.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass

        # Ensure default snippets exist when the feature is enabled
        if self._adv('adv_snippets'):
            try:
                self._ensure_default_snippets()
            except Exception:
                pass

        # Startup lock state (UI is enforced elsewhere)
        try:
            if self._startup_requires_unlock():
                self._unlocked = False
            else:
                self._unlocked = True
        except Exception:
            self._unlocked = True

        # Register/Unregister the template trigger hook
        try:
            self._register_tmplt_trigger()
        except Exception:
            pass

    # ----- App lock (PIN) -----
    def _pin_is_set(self) -> bool:
        try:
            return bool(self.security.get('pin_hash')) and bool(self.security.get('pin_salt'))
        except Exception:
            return False

    def _pin_hash(self, pin: str, salt_b64: str, iters: int) -> str:
        salt = base64.b64decode(salt_b64.encode('utf-8'))
        dk = hashlib.pbkdf2_hmac('sha256', pin.encode('utf-8'), salt, int(iters))
        return dk.hex()

    def _verify_pin_value(self, pin: str) -> bool:
        try:
            if not self._pin_is_set():
                return False
            salt = str(self.security.get('pin_salt') or '')
            iters = int(self.security.get('pin_iters') or 200_000)
            expected = str(self.security.get('pin_hash') or '')
            got = self._pin_hash(pin, salt, iters)
            return secrets.compare_digest(got, expected)
        except Exception:
            return False

    def _save_security(self):
        try:
            safe_json_save(self.security_path, self.security)
        except Exception:
            pass

    # -----------------------------
    # Nuke PIN (emergency local wipe)
    # -----------------------------
    def _nuke_pin_is_set(self) -> bool:
        try:
            salt = str(self.security.get('nuke_pin_salt') or '').strip()
            h = str(self.security.get('nuke_pin_hash') or '').strip()
            return bool(salt and h)
        except Exception:
            return False

    def _verify_nuke_pin_value(self, pin: str) -> bool:
        try:
            if not self._nuke_pin_is_set():
                return False
            salt = str(self.security.get('nuke_pin_salt') or '')
            iters = int(self.security.get('nuke_pin_iters') or 250_000)
            expected = str(self.security.get('nuke_pin_hash') or '')
            got = self._pin_hash(str(pin or '').strip(), salt, iters)
            # constant-time compare
            try:
                return secrets.compare_digest(got, expected)
            except Exception:
                return got == expected
        except Exception:
            return False

    def _set_or_change_nuke_pin_flow(self, parent=None):
        """Set/Change the Nuke PIN (requires current app PIN first)."""
        parent = parent or self

        # Require current PIN (prevents takeover)
        if not self._pin_session_verified:
            cur = simpledialog.askstring(APP_NAME, 'Enter current PIN (required)', parent=parent, show='*')
            if cur is None:
                return
            if not self._verify_pin_value(cur):
                messagebox.showerror(APP_NAME, 'Incorrect PIN.')
                return
            self._pin_session_verified = True

        a = simpledialog.askstring(APP_NAME, 'Set Nuke PIN', parent=parent, show='*')
        if a is None:
            return
        b = simpledialog.askstring(APP_NAME, 'Confirm Nuke PIN', parent=parent, show='*')
        if b is None:
            return
        if a != b:
            messagebox.showerror(APP_NAME, 'Nuke PINs do not match.')
            return
        if len(a.strip()) < 4:
            messagebox.showerror(APP_NAME, 'Nuke PIN must be at least 4 characters.')
            return

        # Stronger defaults; still PBKDF2 via existing helper
        salt = base64.b64encode(secrets.token_bytes(16)).decode('utf-8')
        iters = 250_000
        h = self._pin_hash(a.strip(), salt, iters)
        self.security['nuke_pin_salt'] = salt
        self.security['nuke_pin_iters'] = iters
        self.security['nuke_pin_hash'] = h
        self._save_security()
        messagebox.showinfo(APP_NAME, 'Nuke PIN updated.')

    def _secure_overwrite_file(self, path: Path, passes: int = 1):
        """Best-effort overwrite then delete. Note: cannot guarantee on SSDs/cloud."""
        try:
            if not path.exists() or not path.is_file():
                return
            try:
                size = path.stat().st_size
            except Exception:
                size = None

            if size is None or size <= 0:
                try:
                    path.unlink(missing_ok=True)
                except Exception:
                    pass
                return

            # Overwrite with random bytes (best-effort)
            for _ in range(max(1, int(passes))):
                try:
                    with open(path, 'r+b', buffering=0) as f:
                        remaining = int(size)
                        chunk = 1024 * 1024
                        while remaining > 0:
                            n = chunk if remaining >= chunk else remaining
                            f.write(secrets.token_bytes(n))
                            remaining -= n
                        try:
                            f.flush()
                            os.fsync(f.fileno())
                        except Exception:
                            pass
                except Exception:
                    break

            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass
        except Exception:
            pass

    def _secure_delete_tree(self, root: Path):
        try:
            if not root.exists():
                return
            if root.is_file():
                self._secure_overwrite_file(root, passes=1)
                return
            # Walk files first
            for p in sorted(root.rglob('*'), key=lambda x: len(str(x)), reverse=True):
                try:
                    if p.is_file():
                        self._secure_overwrite_file(p, passes=1)
                    else:
                        try:
                            p.rmdir()
                        except Exception:
                            pass
                except Exception:
                    continue
            try:
                root.rmdir()
            except Exception:
                pass
        except Exception:
            pass

    def _nuke_wipe_everything(self):
        """Wipe all local app data + sync copies, then exit.

        NOTE: This is a destructive action intended for the Nuke PIN path.
        Per requirements, it performs the wipe immediately with no additional prompts.
        """
        try:
            # Stop periodic jobs first
            try:
                self._stop_sync_job()
            except Exception:
                pass
            try:
                if getattr(self, '_idle_after_id', None) is not None:
                    self.after_cancel(self._idle_after_id)
            except Exception:
                pass

            # Clear in-memory state quickly (best-effort)
            try:
                self.history = deque([], maxlen=getattr(self.settings, 'max_history', DEFAULT_MAX_HISTORY))
            except Exception:
                pass
            try:
                self.favorites = []
                self.pins = []
                self.tags = {}
                self.expiry = {}
            except Exception:
                pass

            # Wipe local data directory
            try:
                self._secure_delete_tree(Path(self.data_dir))
            except Exception:
                pass

            # Wipe sync folder copies (known filenames)
            try:
                if self._sync_enabled():
                    folder = Path(str(self.settings.sync_folder)).expanduser()
                    for _local, fname in self._sync_paths():
                        try:
                            self._secure_overwrite_file(folder / fname, passes=1)
                        except Exception:
                            pass
            except Exception:
                pass

        finally:
            # No UI indication for nuke path (silent exit)
            try:
                self.destroy()
            except Exception:
                try:
                    os._exit(0)
                except Exception:
                    pass


    def _set_or_change_pin_flow(self, parent=None):
        parent = parent or self
        # Always require the current PIN before changing it (prevents silent PIN takeover).
        if self._pin_is_set():
            cur = simpledialog.askstring(APP_NAME, 'Enter current PIN', parent=parent, show='*')
            if cur is None:
                return
            if not self._verify_pin_value(cur):
                messagebox.showerror(APP_NAME, 'Incorrect PIN.')
                return
            # Mark session verified after a successful check.
            self._pin_session_verified = True

        a = simpledialog.askstring(APP_NAME, 'Set new PIN', parent=parent, show='*')
        if a is None:
            return
        b = simpledialog.askstring(APP_NAME, 'Confirm new PIN', parent=parent, show='*')
        if b is None:
            return
        if a != b:
            messagebox.showerror(APP_NAME, 'PINs do not match.')
            return
        if len(a.strip()) < 4:
            messagebox.showerror(APP_NAME, 'PIN must be at least 4 characters.')
            return

        salt = base64.b64encode(secrets.token_bytes(16)).decode('utf-8')
        iters = 200_000
        h = self._pin_hash(a.strip(), salt, iters)
        self.security['pin_salt'] = salt
        self.security['pin_iters'] = iters
        self.security['pin_hash'] = h
        self._pin_session_verified = True
        self._save_security()
        messagebox.showinfo(APP_NAME, 'PIN updated.')


    def _prompt_unlock_dialog(self) -> bool:
        """Blocking unlock dialog. Returns True if unlocked, else False (Exit)."""
        dlg = tk.Toplevel(self)
        dlg.title('Unlock')
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        try:
            dlg.attributes('-topmost', True)
        except Exception:
            pass

        frm = ttk.Frame(dlg, padding=18)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text=APP_NAME, font=('Segoe UI', 14, 'bold')).pack(anchor='w')
        ttk.Label(frm, text='Enter your PIN to unlock.').pack(anchor='w', pady=(6, 14))

        pin_var = tk.StringVar(value='')
        ent = ttk.Entry(frm, textvariable=pin_var, show='*', width=24)
        ent.pack(anchor='w')
        ent.focus_set()

        status = tk.StringVar(value='')
        ttk.Label(frm, textvariable=status, foreground='#c00').pack(anchor='w', pady=(8, 0))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(14, 0))

        result = {'ok': False}

        def do_unlock(_evt=None):
            p = pin_var.get() or ''
            # Nuke PIN: emergency wipe (if configured)
            try:
                if self._verify_nuke_pin_value(p):
                    try:
                        pin_var.set('')
                    except Exception:
                        pass
                    try:
                        dlg.destroy()
                    except Exception:
                        pass
                    self._nuke_wipe_everything()
                    return
            except Exception:
                pass

            if self._verify_pin_value(p):
                result['ok'] = True
                # Cache PIN in-memory for this session so encrypted-at-rest stores can be read.
                try:
                    self._session_pin = str(p).strip()
                except Exception:
                    self._session_pin = None
                self._pin_session_verified = True
                try:
                    dlg.destroy()
                except Exception:
                    pass
            else:
                status.set('Incorrect PIN.')

        def do_exit(_evt=None):
            result['ok'] = False
            try:
                dlg.destroy()
            except Exception:
                pass

        ttk.Button(btns, text='Exit', command=do_exit).pack(side=tk.RIGHT)
        ttk.Button(btns, text='Unlock', command=do_unlock).pack(side=tk.RIGHT, padx=(0, 8))

        dlg.protocol('WM_DELETE_WINDOW', do_exit)
        dlg.bind('<Return>', do_unlock)
        dlg.bind('<Escape>', do_exit)

        # Center-ish
        try:
            dlg.update_idletasks()
            w = dlg.winfo_width()
            h = dlg.winfo_height()
            x = max(40, int(self.winfo_screenwidth()/2 - w/2))
            y = max(40, int(self.winfo_screenheight()/2 - h/2))
            dlg.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

        self.wait_window(dlg)
        return bool(result['ok'])

    def _unlock_startup_flow(self):
        """Enforce App Lock on startup. If user cancels, exit the app."""
        try:
            if not (self._adv('adv_app_lock') and self._pin_is_set()):
                self._unlocked = True
                return
            # Hide main window behind lock
            try:
                self.withdraw()
            except Exception:
                pass
            ok = self._prompt_unlock_dialog()
            if ok:
                self._unlocked = True
                self._pin_session_verified = True
                try:
                    self.deiconify()
                except Exception:
                    pass
            else:
                try:
                    self.destroy()
                except Exception:
                    pass
        except Exception:
            pass

    # ----- Start on boot (Windows) -----
    def _set_start_on_boot(self, enabled: bool):
        if winreg is None:
            return
        try:
            run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_SET_VALUE) as k:
                if enabled:
                    # For frozen builds: sys.executable points to Copy2.exe
                    exe = sys.executable
                    if getattr(sys, 'frozen', False):
                        cmd = f'"{exe}"'
                    else:
                        cmd = f'"{exe}" "{os.path.abspath(__file__)}"'
                    winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, cmd)
                else:
                    try:
                        winreg.DeleteValue(k, APP_NAME)
                    except FileNotFoundError:
                        pass
        except Exception:
            pass

    # ----- Snippets / Templates -----
    def _ensure_default_snippets(self):
        if getattr(self, 'snippets', None) is None:
            self.snippets = []
        if isinstance(self.snippets, list) and len(self.snippets) > 0:
            return
        now = now_ts()
        self.snippets = [
            {
                'name': 'Professional — Email Reply',
                'group': 'Professional',
                'body': """Hi {name},

Thanks for your message. {context}

Next steps:
- {next_step_1}
- {next_step_2}

Kind regards,
{signature}""",
                'created_at': now,
                'updated_at': now,
            },
            {
                'name': 'Professional — Meeting Agenda',
                'group': 'Professional',
                'body': """Agenda — {topic}

1) Objective
2) Updates
3) Decisions
4) Action items

Attendees: {attendees}
Time: {time}""",
                'created_at': now,
                'updated_at': now,
            },
            {
                'name': 'Technical — Bug Report',
                'group': 'Technical',
                'body': """Title: {title}

Environment:
- App version: {version}
- OS: {os}

Steps to reproduce:
1) {step1}
2) {step2}

Expected:
{expected}

Actual:
{actual}

Notes / logs:
{notes}""",
                'created_at': now,
                'updated_at': now,
            },
            {
                'name': 'Casual — Quick Reply',
                'group': 'Casual',
                'body': """Hey! {message}

Thanks — {signature}""",
                'created_at': now,
                'updated_at': now,
            },
        ]
        try:
            self._save_snippets()
        except Exception:
            pass

    def _save_snippets(self):
        try:
            self._store_save_json(self.snippets_path, {'templates': getattr(self, 'snippets', [])})
        except Exception:
            pass

    def _paste_text_to_active_app(self, text: str):
        if not isinstance(text, str):
            text = str(text)
        prev = None
        try:
            prev = self._clipboard_get_text()
        except Exception:
            prev = None
        self._clipboard_set_text(text)
        if _kbd is not None:
            try:
                _kbd.send('ctrl+v')
            except Exception:
                pass
        # restore
        if isinstance(prev, str):
            try:
                self._clipboard_set_text(prev)
            except Exception:
                pass

    def _open_snippets_manager(self):
        if not self._adv('adv_snippets'):
            messagebox.showinfo(APP_NAME, 'Enable Snippets under Settings → Advanced Features.')
            return
        self._ensure_default_snippets()

        dlg = tk.Toplevel(self)
        dlg.title('Snippets / Templates')
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry('920x580')
        dlg.minsize(860, 520)

        root = ttk.Frame(dlg, padding=10)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=2)
        root.rowconfigure(1, weight=1)

        ttk.Label(root, text='Search').grid(row=0, column=0, sticky='w')
        q = tk.StringVar(value='')
        ent = ttk.Entry(root, textvariable=q)
        ent.grid(row=0, column=0, sticky='ew', pady=(6, 10), padx=(60, 10))

        lst = tk.Listbox(root, exportselection=False)
        lst.grid(row=1, column=0, sticky='nsew', padx=(0, 10))

        editor = tk.Text(root, wrap='word', undo=True)
        editor.grid(row=1, column=1, sticky='nsew')

        btns = ttk.Frame(root)
        btns.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        btns.columnconfigure(0, weight=1)

        state = {'items': [], 'sel': None}

        def filtered():
            qq = (q.get() or '').strip().lower()
            out = []
            for t in getattr(self, 'snippets', []) or []:
                name = str(t.get('name') or '')
                grp = str(t.get('group') or '')
                body = str(t.get('body') or '')
                if not qq or qq in name.lower() or qq in grp.lower() or qq in body.lower():
                    out.append(t)
            return out

        def refresh():
            items = filtered()
            state['items'] = items
            lst.delete(0, tk.END)
            for t in items:
                name = str(t.get('name') or 'Unnamed')
                grp = str(t.get('group') or '')
                prefix = f"[{grp}] " if grp else ''
                lst.insert(tk.END, prefix + name)

        def on_sel(_=None):
            try:
                i = int(lst.curselection()[0])
            except Exception:
                return
            t = state['items'][i]
            state['sel'] = t
            editor.delete('1.0', tk.END)
            editor.insert('1.0', str(t.get('body') or ''))

        def new_template():
            name = simpledialog.askstring(APP_NAME, 'Template name', parent=dlg)
            if not name:
                return
            grp = simpledialog.askstring(APP_NAME, 'Group (optional)', parent=dlg) or ''
            now = now_ts()
            t = {'name': name.strip(), 'group': grp.strip(), 'body': '', 'created_at': now, 'updated_at': now}
            self.snippets.append(t)
            self._save_snippets()
            refresh()
            # select last
            try:
                lst.selection_clear(0, tk.END)
                lst.selection_set(tk.END)
                lst.see(tk.END)
                on_sel()
            except Exception:
                pass

        def save_template():
            t = state.get('sel')
            if not t:
                return
            t['body'] = editor.get('1.0', tk.END).rstrip("\n")
            t['updated_at'] = now_ts()
            self._save_snippets()
            refresh()

        def delete_template():
            t = state.get('sel')
            if not t:
                return
            if not messagebox.askyesno(APP_NAME, f"Delete template '{t.get('name')}'?"):
                return
            try:
                self.snippets.remove(t)
            except Exception:
                pass
            state['sel'] = None
            editor.delete('1.0', tk.END)
            self._save_snippets()
            refresh()

        def insert_into_app():
            t = state.get('sel')
            if not t:
                return
            self._paste_text_to_active_app(str(t.get('body') or ''))

        def export_templates():
            path = filedialog.asksaveasfilename(title='Export Templates', defaultextension='.json', filetypes=[('JSON','*.json')])
            if not path:
                return
            try:
                Path(path).write_text(json.dumps({'templates': self.snippets}, ensure_ascii=False, indent=2), encoding='utf-8')
            except Exception as e:
                messagebox.showerror(APP_NAME, f"Export failed\n{e}")

        def import_templates():
            path = filedialog.askopenfilename(title='Import Templates', filetypes=[('JSON','*.json')])
            if not path:
                return
            try:
                obj = json.loads(Path(path).read_text(encoding='utf-8'))
                items = []
                if isinstance(obj, dict) and isinstance(obj.get('templates'), list):
                    items = obj.get('templates')
                elif isinstance(obj, list):
                    items = obj
                clean = []
                for t in items:
                    if not isinstance(t, dict):
                        continue
                    name = str(t.get('name') or '').strip()
                    if not name:
                        continue
                    clean.append({
                        'name': name,
                        'group': str(t.get('group') or '').strip(),
                        'body': str(t.get('body') or ''),
                        'created_at': str(t.get('created_at') or now_ts()),
                        'updated_at': now_ts(),
                    })
                if clean:
                    self.snippets.extend(clean)
                    self._save_snippets()
                    refresh()
            except Exception as e:
                messagebox.showerror(APP_NAME, f"Import failed\n{e}")

        ttk.Button(btns, text='New', command=new_template).pack(side=tk.LEFT)
        ttk.Button(btns, text='Save', command=save_template).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text='Delete', command=delete_template).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text='Insert into App', command=insert_into_app).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text='Import', command=import_templates).pack(side=tk.RIGHT)
        ttk.Button(btns, text='Export', command=export_templates).pack(side=tk.RIGHT, padx=(0, 8))

        lst.bind('<<ListboxSelect>>', on_sel)
        q.trace_add('write', lambda *_: refresh())
        refresh()
        ent.focus_set()

    # ----- Template trigger (tmplt) -----
    def _unregister_tmplt_trigger(self):
        try:
            if _kbd is not None and getattr(self, '_tmplt_hook', None) is not None:
                try:
                    _kbd.unhook(self._tmplt_hook)
                except Exception:
                    pass
        except Exception:
            pass
        self._tmplt_hook = None
        self._tmplt_buffer = ''
        self._tmplt_last_ts = 0.0

    def _register_tmplt_trigger(self):
        self._unregister_tmplt_trigger()
        if not (self._adv('adv_tmplt_trigger') and self._adv('adv_snippets')):
            return
        if _kbd is None:
            return
        trigger = str(getattr(self.settings, 'tmplt_trigger_word', 'tmplt') or 'tmplt').strip().lower()
        if not trigger:
            trigger = 'tmplt'

        def handler(e):
            try:
                # reset buffer if idle
                now = time.time()
                if now - float(getattr(self, '_tmplt_last_ts', 0.0)) > 2.0:
                    self._tmplt_buffer = ''
                self._tmplt_last_ts = now

                name = str(getattr(e, 'name', '') or '')
                if len(name) == 1 and name.isprintable():
                    self._tmplt_buffer += name.lower()
                    self._tmplt_buffer = self._tmplt_buffer[-40:]
                elif name in ('space', 'enter', 'tab'):
                    # word boundary; check before adding boundary
                    if self._tmplt_buffer.endswith(trigger):
                        # Remove trigger text in active app (best-effort)
                        try:
                            for _ in range(len(trigger)):
                                _kbd.send('backspace')
                        except Exception:
                            pass
                        # open selector
                        self.after(0, self._open_tmplt_overlay)
                    self._tmplt_buffer = ''
                elif name in ('backspace',):
                    self._tmplt_buffer = self._tmplt_buffer[:-1]
                else:
                    # punctuation boundary
                    if self._tmplt_buffer.endswith(trigger):
                        try:
                            for _ in range(len(trigger)):
                                _kbd.send('backspace')
                        except Exception:
                            pass
                        self.after(0, self._open_tmplt_overlay)
                    self._tmplt_buffer = ''
            except Exception:
                pass

        try:
            self._tmplt_hook = _kbd.on_press(handler)
        except Exception:
            self._tmplt_hook = None

    def _cursor_pos(self):
        try:
            pt = ctypes.wintypes.POINT()  # type: ignore
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))  # type: ignore
            return int(pt.x), int(pt.y)
        except Exception:
            try:
                return int(self.winfo_pointerx()), int(self.winfo_pointery())
            except Exception:
                return 200, 200

    def _open_tmplt_overlay(self):
        if not self._adv('adv_snippets'):
            return
        self._ensure_default_snippets()
        dlg = tk.Toplevel(self)
        dlg.title('Insert Template')
        dlg.transient(self)
        dlg.resizable(False, False)
        try:
            dlg.attributes('-topmost', True)
        except Exception:
            pass

        x, y = self._cursor_pos()
        dlg.geometry(f"420x340+{x}+{y}")

        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text='Search templates').pack(anchor='w')
        q = tk.StringVar(value='')
        ent = ttk.Entry(frm, textvariable=q)
        ent.pack(fill=tk.X, pady=(6, 8))

        lst = tk.Listbox(frm, height=10, exportselection=False)
        lst.pack(fill=tk.BOTH, expand=True)

        hint = ttk.Label(frm, text='Enter to insert • Esc to close', font=('Segoe UI', 9))
        hint.pack(anchor='w', pady=(8, 0))

        state = {'items': []}

        def filtered():
            qq = (q.get() or '').strip().lower()
            out = []
            for t in getattr(self, 'snippets', []) or []:
                name = str(t.get('name') or '')
                grp = str(t.get('group') or '')
                body = str(t.get('body') or '')
                if not qq or qq in name.lower() or qq in grp.lower() or qq in body.lower():
                    out.append(t)
            return out

        def refresh():
            items = filtered()
            state['items'] = items
            lst.delete(0, tk.END)
            for t in items:
                name = str(t.get('name') or 'Unnamed')
                grp = str(t.get('group') or '')
                prefix = f"[{grp}] " if grp else ''
                lst.insert(tk.END, prefix + name)
            if items:
                lst.selection_set(0)

        def do_insert(_evt=None):
            try:
                i = int(lst.curselection()[0])
            except Exception:
                return
            t = state['items'][i]
            self._paste_text_to_active_app(str(t.get('body') or ''))
            try:
                dlg.destroy()
            except Exception:
                pass

        def do_close(_evt=None):
            try:
                dlg.destroy()
            except Exception:
                pass

        q.trace_add('write', lambda *_: refresh())
        lst.bind('<Return>', do_insert)
        lst.bind('<Double-Button-1>', do_insert)
        dlg.bind('<Escape>', do_close)
        ent.bind('<Return>', do_insert)
        refresh()
        ent.focus_set()

    # ----- Images & screenshots -----
    def _save_images_meta(self):
        try:
            self._store_save_json(self.images_meta_path, getattr(self, 'images', []))
        except Exception:
            pass

    def _clipboard_get_image(self):
        if ImageGrab is None:
            return None
        try:
            data = ImageGrab.grabclipboard()
            if data is None:
                return None
            if Image is not None and isinstance(data, Image.Image):
                return data
            return None
        except Exception:
            return None

    def _image_signature(self, img) -> str:
        try:
            import io
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            b = buf.getvalue()
            return hashlib.sha256(b).hexdigest()
        except Exception:
            return ''

    def _add_image_from_pil(self, img):
        if not self._adv('adv_images'):
            return
        try:
            self.images_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        sig = self._image_signature(img)
        if not sig or sig == getattr(self, '_last_image_sig', ''):
            return
        self._last_image_sig = sig

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        rid = secrets.token_hex(4)
        fname = f"img_{ts}_{rid}.png"
        fpath = self.images_dir / fname
        try:
            img.save(str(fpath), format='PNG')
        except Exception:
            return


        # Encrypt image-on-disk if Encrypt-All is enabled (Option A uses PIN)
        try:
            if self._enc_all_enabled() and getattr(self, '_session_pin', None) and Fernet is not None:
                from pathlib import Path as _P
                b = _P(str(fpath)).read_bytes()
                env = self._encrypt_image_bytes(b, self._session_pin, ext='png')
                enc_path = self.images_dir / f"img_{ts}_{rid}.c2img"
                enc_path.write_text(json.dumps(env, ensure_ascii=False, indent=2), encoding='utf-8')
                try:
                    _P(str(fpath)).unlink()
                except Exception:
                    try:
                        os.remove(str(fpath))
                    except Exception:
                        pass
                fpath = enc_path
        except Exception:
            pass

        rec = {
            'id': f"{ts}_{rid}",
            'path': str(fpath),
            'created_at': now_ts(),
        }
        if not isinstance(getattr(self, 'images', None), list):
            self.images = []
        self.images.append(rec)
        self._save_images_meta()

        # If currently in Images view, refresh list
        try:
            if hasattr(self, 'filter_var') and str(self.filter_var.get()) == 'img':
                self._refresh_list(select_last=True)
        except Exception:
            pass

    def _capture_screenshot(self):
        if not self._adv('adv_screenshots'):
            messagebox.showinfo(APP_NAME, 'Enable Screenshots under Settings → Advanced Features.')
            return
        if ImageGrab is None:
            messagebox.showerror(APP_NAME, 'Screenshots require Pillow (PIL).')
            return
        try:
            img = ImageGrab.grab()
            self._add_image_from_pil(img)
            # Switch to Images view so the user can immediately see the result
            try:
                if hasattr(self, 'filter_var'):
                    self.filter_var.set('img')
                self._refresh_list(select_last=True)
            except Exception:
                pass
            self.status_var.set(f"Screenshot captured — {now_ts()}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Screenshot failed\n{e}")

    def _open_images_folder(self):
        try:
            self.images_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            os.startfile(str(self.images_dir))  # type: ignore
        except Exception:
            try:
                webbrowser.open(str(self.images_dir))
            except Exception:
                pass

    def _open_data_folder(self):
        """Open the app's data directory (AppData folder used by Copy2)."""
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            os.startfile(str(self.data_dir))  # type: ignore
        except Exception:
            try:
                webbrowser.open(str(self.data_dir))
            except Exception:
                pass

    def _clipboard_set_image(self, pil_img) -> bool:
        """Copy a PIL image to the Windows clipboard (CF_DIB).

        Falls back to False on non-Windows platforms.
        """
        if pil_img is None:
            return False
        if os.name != 'nt':
            return False
        try:
            import io
            img = pil_img.convert('RGB')
            output = io.BytesIO()
            # BMP includes a 14-byte file header; CF_DF_DIB expects the DIB payload.
            img.save(output, 'BMP')
            data = output.getvalue()[14:]
            output.close()

            GMEM_MOVEABLE = 0x0002
            CF_DIB = 8

            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            if not user32.OpenClipboard(None):
                return False
            try:
                user32.EmptyClipboard()
                hglob = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
                if not hglob:
                    return False
                lp = kernel32.GlobalLock(hglob)
                if not lp:
                    return False
                ctypes.memmove(lp, data, len(data))
                kernel32.GlobalUnlock(hglob)
                if not user32.SetClipboardData(CF_DIB, hglob):
                    return False
                # On success, the clipboard owns the memory handle.
                return True
            finally:
                user32.CloseClipboard()
        except Exception:
            return False

    def _copy_selected_image(self, rec: dict | None) -> bool:
        """Copy selected image to clipboard; fall back to copying file path as text."""
        if not rec:
            return False
        path = str(rec.get('path') or '').strip()
        if not path:
            return False

        try:
            if Image is not None:
                import io
                b = self._load_image_bytes(rec)
                if b is None:
                    raise RuntimeError("no image bytes")
                img = Image.open(io.BytesIO(b))
                if self._clipboard_set_image(img):
                    return True
        except Exception:
            pass

        # Fallback: copy file path
        try:
            return bool(self._clipboard_set_text(path))
        except Exception:
            return False

    def _log_check(self, msg: str):
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
            self.log_update_check.write_text("", encoding="utf-8") if not self.log_update_check.exists() else None
            with self.log_update_check.open("a", encoding="utf-8", errors="ignore") as f:
                f.write(f"{now_ts()}  {msg}\n")
        except Exception:
            pass

    def _log_install(self, msg: str):
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
            self.log_update_install.write_text("", encoding="utf-8") if not self.log_update_install.exists() else None
            with self.log_update_install.open("a", encoding="utf-8", errors="ignore") as f:
                f.write(f"{now_ts()}  {msg}\n")
        except Exception:
            pass

    

    # -----------------------------
    # Data folder + encryption helpers
    # -----------------------------
    _ENC_MAGIC = "__copy2_enc__"
    _ENC_V = 1

    def _open_data_folder(self):
        """Open the app data directory in Explorer."""
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            os.startfile(str(self.data_dir))  # type: ignore
        except Exception:
            try:
                webbrowser.open(str(self.data_dir))
            except Exception:
                pass

    def _enc_all_enabled(self) -> bool:
        """True when 'Encrypt ALL local data files' is active and crypto is available."""
        try:
            if not (getattr(self.settings, 'advanced_features', False) and getattr(self.settings, 'adv_encrypt_all_data', False)):
                return False
            if Fernet is None:
                return False
            return True
        except Exception:
            return False

    def _startup_requires_unlock(self) -> bool:
        """Startup requires unlock when App Lock is on OR local encryption is on, and a PIN exists."""
        try:
            if not getattr(self.settings, 'advanced_features', False):
                return False
            need = bool(getattr(self.settings, 'adv_app_lock', False) or getattr(self.settings, 'adv_encrypt_all_data', False))
            return bool(need and self._pin_is_set())
        except Exception:
            return False

    

    

    # -----------------------------
    # Inactivity auto-lock (UI-only)
    # -----------------------------
    def _bind_inactivity_listeners(self):
        """Bind UI events to reset the inactivity timer.

        Only active when Advanced Features + App Lock are enabled.
        """
        try:
            self.bind_all('<Any-KeyPress>', self._on_user_activity, add='+')
            self.bind_all('<ButtonPress>', self._on_user_activity, add='+')
            self.bind_all('<MouseWheel>', self._on_user_activity, add='+')
            # Additional signals (some widgets don't emit MouseWheel/KeyPress reliably)
            self.bind_all('<Motion>', self._on_user_activity, add='+')
            self.bind_all('<ButtonRelease>', self._on_user_activity, add='+')
            self.bind_all('<KeyRelease>', self._on_user_activity, add='+')
            try:
                self.bind('<FocusIn>', self._on_user_activity, add='+')
            except Exception:
                pass
        except Exception:
            pass

    def _on_user_activity(self, event=None):
        try:
            self._schedule_inactivity_lock()
        except Exception:
            pass

    def _schedule_inactivity_lock(self):
        """(Re)schedule the UI auto-lock timer based on settings."""
        try:
            # Cancel existing timer
            if getattr(self, '_idle_after_id', None) is not None:
                try:
                    self.after_cancel(self._idle_after_id)
                except Exception:
                    pass
                self._idle_after_id = None

            # Only when App Lock is enabled
            if not (getattr(self.settings, 'advanced_features', False) and getattr(self.settings, 'adv_app_lock', False)):
                return

            minutes = int(getattr(self.settings, 'lock_timeout_minutes', 0) or 0)
            if minutes <= 0:
                return

            # If currently locked, no need to schedule
            if not getattr(self, '_unlocked', True):
                return

            ms = minutes * 60 * 1000
            self._idle_after_id = self.after(ms, self._auto_lock_due_to_inactivity)
        except Exception:
            self._idle_after_id = None

    def _auto_lock_due_to_inactivity(self):
        """Lock UI after inactivity.

        This must lock the *entire* UI surface, including any open dialogs (Settings, Quick Paste, etc.).
        Clipboard engine continues running in the background.
        """
        try:
            # If app lock not enabled anymore, do nothing
            if not (getattr(self.settings, 'advanced_features', False) and getattr(self.settings, 'adv_app_lock', False)):
                return

            # Mark locked first (so new dialogs/features are gated)
            self._unlocked = False

            # Ensure secondary windows cannot remain interactive on top of the lock overlay
            try:
                self._close_non_main_windows_on_lock()
            except Exception:
                pass

            # Show overlay on the root window
            self._show_lock_overlay(reason='Locked')

            # Keep the main window above any remaining windows
            try:
                self.lift()
                self.focus_force()
            except Exception:
                pass
        except Exception:
            pass

    def _close_non_main_windows_on_lock(self):
        """Hide/disable any open Toplevel windows so the lock applies globally.

        Uses a robust enumeration of Tk toplevel children so dialogs created with different masters
        are still captured.
        """
        try:
            self._withdrawn_on_lock = []
        except Exception:
            self._withdrawn_on_lock = []

        # Enumerate all toplevel children under the Tk root ('.')
        names = []
        try:
            names = list(self.tk.call('winfo', 'children', '.'))
        except Exception:
            names = []

        for n in names:
            try:
                w = self.nametowidget(n)
            except Exception:
                continue
            try:
                # Skip self (the main root)
                if w is self:
                    continue
            except Exception:
                pass
            try:
                if isinstance(w, (tk.Toplevel, tk.Tk)) and w.winfo_exists():
                    try:
                        st = str(w.state() or '')
                    except Exception:
                        st = ''
                    if st == 'withdrawn':
                        continue
                    try:
                        self._withdrawn_on_lock.append(w)
                    except Exception:
                        pass
                    try:
                        w.withdraw()
                    except Exception:
                        try:
                            w.attributes('-disabled', True)
                        except Exception:
                            pass
            except Exception:
                continue

    def _restore_windows_after_unlock(self):
        """Restore any windows that were withdrawn when the app locked."""
        ws = getattr(self, '_withdrawn_on_lock', None) or []
        try:
            self._withdrawn_on_lock = []
        except Exception:
            pass
        for w in ws:
            try:
                if w is not None and w.winfo_exists():
                    try:
                        w.attributes('-disabled', False)
                    except Exception:
                        pass
                    try:
                        w.deiconify()
                    except Exception:
                        pass
            except Exception:
                pass

# -----------------------------
    # UI-only Lock Overlay (no modal dialogs, no withdraw/iconify) (no modal dialogs, no withdraw/iconify)
    # -----------------------------
    def _require_unlocked(self, feature_name: str = "") -> bool:
        """Gate UI-only features while locked. Engine (clipboard polling) must continue."""
        try:
            if getattr(self, '_unlocked', True):
                return True
        except Exception:
            return True
        try:
            self._show_lock_overlay(reason=(feature_name or "Locked"))
        except Exception:
            pass
        return False

    def _install_lock_blocker_bindings_once(self):
        """Bind global input blockers so the entire app becomes non-interactive while locked.

        This protects against cases where a widget/dialog remains above the overlay or where geometry results
        in partial coverage. The blocker allows events only within the lock overlay subtree.
        """
        if getattr(self, '_lock_blocker_bound', False):
            return
        self._lock_blocker_bound = True

        def blocker(event):
            # Only block when locked
            try:
                if getattr(self, '_unlocked', True):
                    return
            except Exception:
                pass

            ov = getattr(self, '_lock_overlay', None)
            if ov is None:
                return "break"

            try:
                w = event.widget
                # Allow interactions inside the lock overlay
                if w == ov or str(w).startswith(str(ov)):
                    return
            except Exception:
                pass

            return "break"

        self._lock_blocker_fn = blocker

        # Broad set of common UI inputs
        seqs = [
            '<ButtonPress>', '<ButtonRelease>',
            '<Button-1>', '<Button-2>', '<Button-3>',
            '<Double-Button-1>',
            '<KeyPress>', '<KeyRelease>',
            '<MouseWheel>',
        ]
        for s in seqs:
            try:
                self.bind_all(s, blocker, add='+')
            except Exception:
                pass

    def _apply_startup_lock_overlay(self):
        """Apply startup lock if App Lock and/or Encrypt-All requires it.

        This does not block app launch and does not stop clipboard polling.
        """
        try:
            if self._startup_requires_unlock():
                self._unlocked = False
                self._show_lock_overlay(reason="Locked")
            else:
                self._unlocked = True
                self._hide_lock_overlay()

            # Start/refresh inactivity timer (if enabled)
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
        except Exception:
            # Fail open visually, but do not break launch
            try:
                self._hide_lock_overlay()
            except Exception:
                pass

    def _ensure_lock_overlay(self):
        if getattr(self, '_lock_overlay', None) is not None:
            return

        # Use plain Tk widgets for maximum compatibility (works under ttkbootstrap too)
        ov = tk.Frame(self)
        try:
            # If ttkbootstrap is active, use theme colors when possible
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is not None:
                ov.configure(bg=colors.bg)
        except Exception:
            pass

        # Center card
        card = tk.Frame(ov)
        try:
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is not None:
                card.configure(bg=colors.bg)
        except Exception:
            pass
        card.place(relx=0.5, rely=0.5, anchor='center')

        # Branding
        # Icon above app name (best-effort)
        try:
            ico = self._get_lock_logo_photo()
        except Exception:
            ico = None
        if ico is not None:
            try:
                icon_lbl = tk.Label(card, image=ico)
                icon_lbl.image = ico  # keep ref
                try:
                    colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
                    if colors is not None:
                        icon_lbl.configure(bg=colors.bg)
                except Exception:
                    pass
                icon_lbl.pack(padx=24, pady=(12, 6))
            except Exception:
                pass

        title = tk.Label(card, text="Copy 2.0", font=("Segoe UI", 22, "bold"))
        subtitle = tk.Label(card, text="Enter PIN to unlock", font=("Segoe UI", 11))
        try:
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is not None:
                title.configure(bg=colors.bg, fg=colors.fg)
                subtitle.configure(bg=colors.bg, fg=colors.fg)
        except Exception:
            pass
        title.pack(padx=24, pady=(0, 4) if ico is not None else (14, 4))
        subtitle.pack(padx=24, pady=(0, 10))

        # PIN entry
        pin_var = tk.StringVar(value="")
        ent = tk.Entry(card, textvariable=pin_var, show='*', width=28, font=("Segoe UI", 12))
        ent.pack(padx=24, pady=(0, 10))

        # Error line
        err_var = tk.StringVar(value="")
        err = tk.Label(card, textvariable=err_var, font=("Segoe UI", 10))
        try:
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is not None:
                err.configure(bg=colors.bg, fg=colors.danger)
        except Exception:
            pass
        err.pack(padx=24, pady=(0, 8))

        # Unlock button (no Cancel)
        btn = tk.Button(card, text="Unlock", command=lambda: self._attempt_unlock_from_overlay(pin_var, err_var))
        btn.pack(padx=24, pady=(6, 18), fill='x')

        # Bind Enter to unlock
        try:
            ent.bind('<Return>', lambda e: (self._attempt_unlock_from_overlay(pin_var, err_var), 'break'))
        except Exception:
            pass

        # Ensure global lock input blocker is installed (one-time)
        try:
            self._install_lock_blocker_bindings_once()
        except Exception:
            pass

        self._lock_overlay = ov
        self._lock_pin_var = pin_var
        self._lock_err_var = err_var
        self._lock_entry = ent

    def _show_lock_overlay(self, reason: str = "Locked"):
        self._ensure_lock_overlay()
        try:
            if getattr(self, '_lock_err_var', None) is not None:
                self._lock_err_var.set("")
        except Exception:
            pass


        # Always clear the PIN field when showing the lock screen
        try:
            if getattr(self, '_lock_pin_var', None) is not None:
                self._lock_pin_var.set('')
        except Exception:
            pass
        # Cover the full client area; blocks interaction with underlying UI
        try:
            self._lock_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        except Exception:
            try:
                self._lock_overlay.pack(fill='both', expand=True)
            except Exception:
                pass

        # Ensure overlay is on top
        try:
            self._lock_overlay.lift()
            self._lock_overlay.tkraise()
        except Exception:
            pass

        # Focus PIN field
        try:
            self._lock_entry.focus_set()
            self._lock_entry.icursor('end')
        except Exception:
            pass

    def _hide_lock_overlay(self):
        ov = getattr(self, '_lock_overlay', None)
        if ov is None:
            return
        try:
            ov.place_forget()
        except Exception:
            pass
        try:
            ov.pack_forget()
        except Exception:
            pass

        # Clear PIN/error fields so they are never retained between locks
        try:
            if getattr(self, '_lock_pin_var', None) is not None:
                self._lock_pin_var.set('')
        except Exception:
            pass
        try:
            if getattr(self, '_lock_err_var', None) is not None:
                self._lock_err_var.set('')
        except Exception:
            pass
        # Restore any dialogs that were withdrawn during lock
        try:
            self._restore_windows_after_unlock()
        except Exception:
            pass


    def _attempt_unlock_from_overlay(self, pin_var: tk.StringVar, err_var: tk.StringVar):
        pin = ''
        try:
            pin = str(pin_var.get() or '').strip()
        except Exception:
            pin = ''

        if not pin:
            try:
                err_var.set('PIN required')
            except Exception:
                pass
            try:
                self._lock_entry.focus_set()
            except Exception:
                pass
            return

        # Nuke PIN: emergency wipe (if configured)
        try:
            if self._verify_nuke_pin_value(pin):
                # Clear entry quickly, then wipe
                try:
                    pin_var.set('')
                except Exception:
                    pass
                self._nuke_wipe_everything()
                return
        except Exception:
            pass

        if not self._verify_pin_value(pin):
            try:
                err_var.set('Incorrect PIN')
                pin_var.set('')
            except Exception:
                pass
            try:
                self._lock_entry.focus_set()
            except Exception:
                pass
            return

        # Success
        try:
            self._session_pin = pin
        except Exception:
            pass
        try:
            self._unlocked = True
            self._pin_session_verified = True
        except Exception:
            pass

        # If encrypted stores were deferred, load them now
        try:
            if not bool(getattr(self, '_stores_loaded', True)):
                self._load_persisted_state()
                self._stores_loaded = True

                # Merge any items captured while locked
                pending = list(getattr(self, '_captured_while_locked', []) or [])
                if pending:
                    for t in pending:
                        try:
                            if isinstance(t, str) and t and t not in self.history:
                                self.history.append(t)
                        except Exception:
                            pass
                    try:
                        self._captured_while_locked.clear()
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            self._refresh_list(select_last=True)
        except Exception:
            pass

        try:
            self._hide_lock_overlay()
        except Exception:
            pass

        # Start inactivity timer now that UI is unlocked
        try:
            self._schedule_inactivity_lock()
        except Exception:
            pass

    def _kdf_fernet_key(self, pin: str, salt: bytes, iters: int = 200_000) -> bytes:
        import base64
        try:
            if PBKDF2HMAC is None or hashes is None:
                raise RuntimeError("PBKDF2 unavailable")
            kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iters)
            key = kdf.derive(pin.encode('utf-8'))
            return base64.urlsafe_b64encode(key)
        except Exception:
            # Last-resort fallback (weak): deterministic SHA256
            import hashlib
            key = hashlib.sha256((pin + salt.hex()).encode('utf-8')).digest()[:32]
            return base64.urlsafe_b64encode(key)

    def _encrypt_json_obj(self, obj: object, pin: str) -> dict:
        import base64
        salt = secrets.token_bytes(16)
        iters = 200_000
        key = self._kdf_fernet_key(pin, salt, iters)
        f = Fernet(key)
        payload = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        token = f.encrypt(payload)
        return {
            self._ENC_MAGIC: self._ENC_V,
            "salt_b64": base64.b64encode(salt).decode('ascii'),
            "iters": iters,
            "token": token.decode('ascii'),
        }

    def _decrypt_json_obj(self, env: dict, pin: str):
        import base64
        salt = base64.b64decode(str(env.get('salt_b64') or ''))
        iters = int(env.get('iters') or 200_000)
        token = str(env.get('token') or '').encode('ascii')
        key = self._kdf_fernet_key(pin, salt, iters)
        f = Fernet(key)
        raw = f.decrypt(token)
        return json.loads(raw.decode('utf-8'))

    def _encrypt_image_bytes(self, blob: bytes, pin: str, ext: str = 'png') -> dict:
        """Encrypt raw image bytes into a Copy2 envelope (stored on disk as JSON)."""
        import base64
        try:
            data_b64 = base64.b64encode(blob).decode('ascii')
        except Exception:
            data_b64 = ''
        return self._encrypt_json_obj({"data_b64": data_b64, "ext": ext}, pin)

    def _store_load_json(self, p: Path, default):
        """Load JSON, supporting Copy2 encrypted envelopes when encrypt-all is enabled."""
        try:
            if not p.exists():
                return default
            txt = p.read_text(encoding='utf-8', errors='ignore')
            data = json.loads(txt)
        except Exception:
            return default

        # Encrypted envelope
        try:
            if isinstance(data, dict) and data.get(self._ENC_MAGIC) == self._ENC_V:
                pin = getattr(self, '_session_pin', None)
                if not pin:
                    return default
                try:
                    return self._decrypt_json_obj(data, pin)
                except Exception:
                    return default
        except Exception:
            pass

        return data if data is not None else default

    def _store_save_json(self, p: Path, obj: object):
        """Save JSON, encrypting when encrypt-all is enabled and a session PIN is present."""
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        if self._enc_all_enabled():
            pin = getattr(self, '_session_pin', None)
            if pin:
                try:
                    env = self._encrypt_json_obj(obj, pin)
                    p.write_text(json.dumps(env, ensure_ascii=False, indent=2), encoding='utf-8')
                    return
                except Exception:
                    pass

        # Plaintext fallback
        try:
            p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')
        except Exception:
            pass

    def _load_image_bytes(self, rec: dict) -> bytes | None:
        """Load raw image bytes, supporting encrypted .c2img files."""
        try:
            path = str(rec.get('path') or '').strip()
            if not path or not os.path.exists(path):
                return None
            if path.lower().endswith('.c2img'):
                if Fernet is None:
                    return None
                pin = getattr(self, '_session_pin', None)
                if not pin:
                    return None
                raw_env = json.loads(Path(path).read_text(encoding='utf-8', errors='ignore'))
                if not (isinstance(raw_env, dict) and raw_env.get(self._ENC_MAGIC) == self._ENC_V):
                    return None
                data = self._decrypt_json_obj(raw_env, pin)
                if isinstance(data, dict) and 'data_b64' in data:
                    import base64
                    return base64.b64decode(data['data_b64'])
                return None
            # Plain image
            return Path(path).read_bytes()
        except Exception:
            return None

    def _migrate_image_files_for_encrypt_all(self, encrypt: bool):
        """Convert existing image files between plain PNG/JPG and encrypted .c2img files.

        This is best-effort. It updates `self.images` records in-place.
        """
        try:
            if not getattr(self, 'images', None):
                return
            if Fernet is None:
                return
            pin = getattr(self, '_session_pin', None)
            if not pin:
                return

            for rec in list(self.images):
                try:
                    path = str(rec.get('path') or '').strip()
                    if not path:
                        continue
                    p = Path(path)
                    if not p.exists():
                        continue

                    if encrypt:
                        if p.suffix.lower() == '.c2img':
                            rec['enc'] = True
                            continue
                        raw = p.read_bytes()
                        ext = (rec.get('ext') or p.suffix.lstrip('.').lower() or 'png')
                        env = self._encrypt_image_bytes(raw, pin, ext=str(ext))
                        outp = p.with_suffix('.c2img')
                        outp.write_text(json.dumps(env, ensure_ascii=False, indent=2), encoding='utf-8')
                        try:
                            p.unlink(missing_ok=True)
                        except Exception:
                            pass
                        rec['path'] = str(outp)
                        rec['enc'] = True
                        rec['ext'] = ext
                    else:
                        # Decrypt
                        if p.suffix.lower() != '.c2img':
                            rec['enc'] = False
                            continue
                        raw_env = json.loads(p.read_text(encoding='utf-8', errors='ignore'))
                        if not (isinstance(raw_env, dict) and raw_env.get(self._ENC_MAGIC) == self._ENC_V):
                            continue
                        data = self._decrypt_json_obj(raw_env, pin)
                        if not isinstance(data, dict) or 'data_b64' not in data:
                            continue
                        ext = str(data.get('ext') or rec.get('ext') or 'png')
                        raw = base64.b64decode(str(data.get('data_b64') or ''))
                        outp = p.with_suffix('.' + ext)
                        outp.write_bytes(raw)
                        try:
                            p.unlink(missing_ok=True)
                        except Exception:
                            pass
                        rec['path'] = str(outp)
                        rec['enc'] = False
                        rec['ext'] = ext
                except Exception:
                    continue
        except Exception:
            return

    def _image_key_for_rec(self, rec: dict, fallback_i: int = 0) -> str:
        rid = str(rec.get('id') or '').strip()
        if rid:
            return f"IMG::{rid}"
        p = str(rec.get('path') or '').strip()
        base = os.path.basename(p) if p else str(fallback_i)
        return f"IMG::{base}"

    def _enforce_startup_security(self) -> bool:
        """Deprecated: startup security is now enforced via a UI-only lock overlay.

        This method is kept for backward compatibility with older init flows.
        It must never hide/withdraw the root or block the clipboard engine.
        """
        try:
            self.after(0, self._apply_startup_lock_overlay)

            # Start inactivity timer (only does something if enabled)
            self.after(0, self._schedule_inactivity_lock)
        except Exception:
            pass
        return True

    def _load_persisted_state(self):
        """(Re)load persisted state using store helpers (supports encrypted stores)."""
        # History
        try:
            hist = self._store_load_json(self.history_path, [])
            self.history = deque([x for x in (hist or []) if isinstance(x, str) and x.strip()], maxlen=self.settings.max_history)
        except Exception:
            pass
        # Lists/dicts
        try:
            self.favorites = list(self._store_load_json(self.favs_path, []))
        except Exception:
            self.favorites = []
        try:
            self.pins = list(self._store_load_json(self.pins_path, []))
        except Exception:
            self.pins = []
        try:
            self.tags = self._store_load_json(self.tags_path, {}) or {}
        except Exception:
            self.tags = {}
        try:
            self.tag_colors = self._store_load_json(self.tag_colors_path, {}) or {}
        except Exception:
            self.tag_colors = {}
        try:
            self.expiry = self._store_load_json(self.expiry_path, {}) or {}
        except Exception:
            self.expiry = {}
        try:
            fmts = self._store_load_json(self.formats_path, {})
            self.clip_formats = fmts if isinstance(fmts, dict) else {}
        except Exception:
            self.clip_formats = {}
        try:
            meta = self._store_load_json(self.images_meta_path, [])
            self.images = meta if isinstance(meta, list) else []
        except Exception:
            self.images = []
        try:
            sn = self._store_load_json(self.snippets_path, {'templates': []})
            self.snippets = sn.get('templates', []) if isinstance(sn, dict) else []
        except Exception:
            self.snippets = []

    def _persist(self):
        """Persist settings and state. When Encrypt-All is enabled, sensitive stores are encrypted."""
        # Always persist settings (plaintext; required to know whether encryption/lock are enabled).
        safe_json_save(self.settings_path, asdict(self.settings))

        # Security data must remain plaintext (PIN verification depends on it).
        try:
            safe_json_save(self.security_path, getattr(self, 'security', {}))
        except Exception:
            pass

        # If Encrypt-All is enabled, never write store files unless a verified session PIN is present.
        # This prevents overwriting encrypted JSON envelopes with plaintext defaults while locked.
        try:
            if self._enc_all_enabled() and not getattr(self, '_session_pin', None):
                return
        except Exception:
            pass

        # Advanced feature stores (encrypt-all applies here too)
        try:
            self._store_save_json(self.snippets_path, {'templates': getattr(self, 'snippets', [])})
            self._store_save_json(self.images_meta_path, getattr(self, 'images', []))
        except Exception:
            pass

        if self.settings.session_only:
            return

        self._store_save_json(self.history_path, list(self.history))
        self._store_save_json(self.favs_path, list(self.favorites))
        self._store_save_json(self.pins_path, list(self.pins))
        self._store_save_json(self.tags_path, self.tags)
        self._store_save_json(self.tag_colors_path, getattr(self, "tag_colors", {}))
        self._store_save_json(self.expiry_path, self.expiry)
        # Rich clipboard formats (HTML/RTF)
        try:
            self._store_save_json(self.formats_path, getattr(self, 'clip_formats', {}) or {})
        except Exception:
            pass


    # -----------------------------
    # Status bar helpers
    # -----------------------------
    def _status_base_line(self) -> str:
        return (
            f"v{APP_VERSION}   Items: {len(self.history)}   Favorites: {len(self.favorites)}   Pins: {len(self.pins)}   Data: {self.data_dir}"
        )

    def _update_status_bar(self, note: str | None = None):
        if note is not None:
            try:
                self._status_note = str(note).strip()
            except Exception:
                self._status_note = ""
        base = self._status_base_line()
        n = getattr(self, '_status_note', '') or ''
        if n:
            base = base + "   |   " + n
        try:
            self.status_var.set(base)
        except Exception:
            pass

    def _set_status_note(self, note: str):
        self._update_status_bar(note)

    def _apply_theme_to_tk_widgets(self):
        """Apply current ttkbootstrap theme colors to non-ttk widgets (Listbox/Text)."""
        try:
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is None:
                return

            if hasattr(self, 'listbox') and getattr(self, 'listbox', None) is not None:
                try:
                    self.listbox.configure(
                        bg=colors.bg,
                        fg=colors.fg,
                        selectbackground=colors.primary,
                        selectforeground=(colors.light if hasattr(colors, 'light') else colors.fg),
                    )
                except Exception:
                    pass

            if hasattr(self, 'preview') and getattr(self, 'preview', None) is not None:
                try:
                    self.preview.configure(
                        bg=colors.bg,
                        fg=colors.fg,
                        insertbackground=colors.fg,
                    )
                except Exception:
                    pass
        except Exception:
            pass

    # -----------------------------
    # Clipboard helpers
    # -----------------------------
    def _clipboard_get_text(self) -> str:
        """Best-effort read of text clipboard. Returns '' if non-text or unavailable."""
        # Prefer pyperclip if present (works even when app not focused).
        try:
            if pyperclip is not None:
                t = pyperclip.paste()
                if isinstance(t, str):
                    return t
        except Exception:
            pass

        # Fallback to Tk clipboard APIs
        try:
            t = self.clipboard_get()
            return t if isinstance(t, str) else ''
        except Exception:
            return ''

    def _clipboard_set_text(self, text_value: str) -> bool:
        """Best-effort set of text clipboard. Returns True on success."""
        t = text_value if isinstance(text_value, str) else str(text_value)
        # Prefer pyperclip if present
        try:
            if pyperclip is not None:
                pyperclip.copy(t)
                return True
        except Exception:
            pass

        # Fallback to Tk clipboard APIs
        try:
            self.clipboard_clear()
            self.clipboard_append(t)
            return True
        except Exception:
            return False

    def _clipboard_get_rich_payload(self) -> dict:
        """Return a payload containing text plus optional HTML/RTF clipboard formats.

        This is Windows-only best-effort. On non-Windows platforms, returns text only.
        """
        out = {"text": "", "html": None, "rtf": None}
        try:
            out["text"] = self._clipboard_get_text()
        except Exception:
            out["text"] = ""

        try:
            if os.name != 'nt':
                return out
        except Exception:
            return out

        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            CF_UNICODETEXT = 13
            fmt_html = user32.RegisterClipboardFormatW("HTML Format")
            fmt_rtf = user32.RegisterClipboardFormatW("Rich Text Format")

            if not user32.OpenClipboard(None):
                return out
            try:
                # HTML
                try:
                    if fmt_html:
                        h = user32.GetClipboardData(fmt_html)
                        if h:
                            p = kernel32.GlobalLock(h)
                            if p:
                                try:
                                    sz = kernel32.GlobalSize(h)
                                    out["html"] = ctypes.string_at(p, sz)
                                finally:
                                    kernel32.GlobalUnlock(h)
                except Exception:
                    pass

                # RTF
                try:
                    if fmt_rtf:
                        h = user32.GetClipboardData(fmt_rtf)
                        if h:
                            p = kernel32.GlobalLock(h)
                            if p:
                                try:
                                    sz = kernel32.GlobalSize(h)
                                    out["rtf"] = ctypes.string_at(p, sz)
                                finally:
                                    kernel32.GlobalUnlock(h)
                except Exception:
                    pass

                # If Unicode text was not available via pyperclip/Tk (rare), attempt it here
                try:
                    if not out.get("text"):
                        htxt = user32.GetClipboardData(CF_UNICODETEXT)
                        if htxt:
                            p = kernel32.GlobalLock(htxt)
                            if p:
                                try:
                                    out["text"] = ctypes.wstring_at(p)
                                finally:
                                    kernel32.GlobalUnlock(htxt)
                except Exception:
                    pass
            finally:
                try:
                    user32.CloseClipboard()
                except Exception:
                    pass
        except Exception:
            return out
        return out

    def _clipboard_set_rich_text(self, text_value: str) -> bool:
        """Set clipboard using the stored HTML/RTF formats for this text when available.

        Falls back to plain text when:
        - not on Windows
        - no stored formats exist for the given text
        - clipboard API fails
        """
        t = text_value if isinstance(text_value, str) else str(text_value)
        try:
            if os.name != 'nt':
                return self._clipboard_set_text(t)
        except Exception:
            return self._clipboard_set_text(t)

        rec = None
        try:
            rec = (getattr(self, 'clip_formats', {}) or {}).get(t)
        except Exception:
            rec = None
        if not isinstance(rec, dict):
            return self._clipboard_set_text(t)

        html_b64 = rec.get('html_b64')
        rtf_b64 = rec.get('rtf_b64')
        want_any = bool(isinstance(html_b64, str) and html_b64) or bool(isinstance(rtf_b64, str) and rtf_b64)
        if not want_any:
            return self._clipboard_set_text(t)

        try:
            import base64
            html = base64.b64decode(html_b64.encode('ascii')) if isinstance(html_b64, str) and html_b64 else None
            rtf = base64.b64decode(rtf_b64.encode('ascii')) if isinstance(rtf_b64, str) and rtf_b64 else None
        except Exception:
            html = rtf = None
        if html is None and rtf is None:
            return self._clipboard_set_text(t)

        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            GMEM_MOVEABLE = 0x0002
            CF_UNICODETEXT = 13
            fmt_html = user32.RegisterClipboardFormatW("HTML Format")
            fmt_rtf = user32.RegisterClipboardFormatW("Rich Text Format")

            if not user32.OpenClipboard(None):
                return self._clipboard_set_text(t)
            try:
                user32.EmptyClipboard()

                def _set_bytes(fmt, data: bytes):
                    if not fmt or not data:
                        return
                    h = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
                    if not h:
                        return
                    p = kernel32.GlobalLock(h)
                    if not p:
                        kernel32.GlobalFree(h)
                        return
                    try:
                        ctypes.memmove(p, data, len(data))
                    finally:
                        kernel32.GlobalUnlock(h)
                    user32.SetClipboardData(fmt, h)

                # Text (always)
                try:
                    data_u16 = (t + "\0").encode('utf-16le')
                    _set_bytes(CF_UNICODETEXT, data_u16)
                except Exception:
                    pass

                # Rich formats (optional)
                try:
                    if isinstance(html, (bytes, bytearray)) and html:
                        _set_bytes(fmt_html, bytes(html))
                except Exception:
                    pass
                try:
                    if isinstance(rtf, (bytes, bytearray)) and rtf:
                        _set_bytes(fmt_rtf, bytes(rtf))
                except Exception:
                    pass
            finally:
                try:
                    user32.CloseClipboard()
                except Exception:
                    pass
            return True
        except Exception:
            return self._clipboard_set_text(t)


    # -----------------------------
    # Global hotkeys (optional) & sync
    # -----------------------------
    def _unregister_global_hotkeys(self):
        try:
            handles = getattr(self, '_global_hotkey_handles', []) or []
            if _kbd is not None:
                for h in handles:
                    try:
                        _kbd.remove_hotkey(h)
                    except Exception:
                        pass
            self._global_hotkey_handles = []
        except Exception:
            self._global_hotkey_handles = []

    def _register_global_hotkeys(self):
        # Register Windows global hotkeys if enabled
        self._unregister_global_hotkeys()
        if not getattr(self.settings, 'enable_global_hotkeys', False):
            return
        if _kbd is None:
            # Do not hard fail; user can still use the app normally
            try:
                self.status_var.set(f"Global hotkeys require the 'keyboard' package — {now_ts()}")
            except Exception:
                pass
            return

        def safe_call(fn):
            def _wrap():
                try:
                    # Run on UI thread
                    self.after(0, fn)
                except Exception:
                    pass
            return _wrap

        handles = []
        try:
            hk_qp = str(getattr(self.settings, 'hotkey_quick_paste', 'ctrl+alt+v') or '').strip()
            hk_last = str(getattr(self.settings, 'hotkey_paste_last', 'ctrl+alt+shift+v') or '').strip()
            hk_pause = str(getattr(self.settings, 'hotkey_toggle_pause', 'ctrl+alt+p') or '').strip()
            if hk_qp:
                handles.append(_kbd.add_hotkey(hk_qp, safe_call(self._open_quick_paste)))
            if hk_last:
                handles.append(_kbd.add_hotkey(hk_last, safe_call(self._paste_last_hotkey)))
            if hk_pause:
                handles.append(_kbd.add_hotkey(hk_pause, safe_call(self._toggle_pause)))
        except Exception:
            pass

        self._global_hotkey_handles = handles

    def _paste_last_hotkey(self):
        # Copy newest history item to clipboard and paste (if possible)
        if not self.history:
            return
        item = list(self.history)[-1]
        if not self._clipboard_set_rich_text(item):
            return

        if _kbd is None:
            # fallback: copied only
            try:
                self.status_var.set(f"Copied latest item (install 'keyboard' to paste) — {now_ts()}")
            except Exception:
                pass
            return

        try:
            # Give focus back to previous app and paste
            self.after(120, lambda: _kbd.send('ctrl+v'))
        except Exception:
            pass

    # ----- Sync (folder-based; suitable for Dropbox/OneDrive/Google Drive) -----
    def _sync_enabled(self) -> bool:
        return bool(getattr(self.settings, 'sync_enabled', False)) and bool(str(getattr(self.settings, 'sync_folder', '') or '').strip())

    def _sync_paths(self):
        # local_path, filename
        return [
            (self.history_path, 'history.json'),
            (self.favs_path, 'favorites.json'),
            (self.pins_path, 'pins.json'),
            (self.tags_path, 'tags.json'),
            (self.expiry_path, 'expiry.json'),
            (self.formats_path, 'formats.json'),
        ]

    def _sync_now(self):
        if not self._sync_enabled():
            messagebox.showinfo(APP_NAME, "Sync is not enabled.\\n\\nEnable it in Settings and choose a sync folder (e.g., a Dropbox/OneDrive directory).")
            return
        self._run_sync_cycle(show_status=True)

    def _start_sync_job(self):
        # periodic sync
        try:
            if getattr(self, '_sync_job', None) is not None:
                self.after_cancel(self._sync_job)
        except Exception:
            pass
        self._sync_job = None

        if not self._sync_enabled():
            return

        interval = int(getattr(self.settings, 'sync_interval_sec', 10) or 10)
        interval = max(5, min(300, interval))

        def tick():
            try:
                self._run_sync_cycle(show_status=False)
            finally:
                try:
                    self._sync_job = self.after(interval * 1000, tick)
                except Exception:
                    self._sync_job = None

        try:
            self._sync_job = self.after(interval * 1000, tick)
        except Exception:
            self._sync_job = None

    def _stop_sync_job(self):
        try:
            if getattr(self, '_sync_job', None) is not None:
                self.after_cancel(self._sync_job)
        except Exception:
            pass
        self._sync_job = None

    def _run_sync_cycle(self, show_status: bool = False):
        if not self._sync_enabled():
            return
        folder = Path(str(self.settings.sync_folder)).expanduser()
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception:
            return

        # Pull then push
        changed = self._sync_pull(folder)
        pushed = self._sync_push(folder)
        if show_status:
            msg = f"Sync complete. Pulled={changed}  Pushed={pushed}"
            try:
                self.status_var.set(f"{msg} — {now_ts()}")
            except Exception:
                pass

    def _sync_pull(self, folder: Path) -> int:
        # Merge remote files into local if remote changed
        changed = 0
        last = getattr(self, '_sync_seen_mtimes', {}) or {}

        for local_path, fname in self._sync_paths():
            rpath = folder / fname
            if not rpath.exists():
                continue
            try:
                r_mtime = rpath.stat().st_mtime
            except Exception:
                continue
            if last.get(fname) is not None and r_mtime <= last.get(fname):
                continue
            # merge based on file type
            try:
                data = safe_json_load(rpath, None)
            except Exception:
                data = None

            if data is None:
                last[fname] = r_mtime
                continue

            if fname == 'history.json' and isinstance(data, list):
                remote = [str(x) for x in data if isinstance(x, (str, int, float))]
                local = list(self.history)
                merged = local + [x for x in remote if x not in local]
                self.history = deque(merged, maxlen=self.settings.max_history)
                self._prune_preserving_favorites()
                changed += 1
            elif fname == 'favorites.json' and isinstance(data, list):
                remote = [str(x) for x in data if isinstance(x, (str, int, float))]
                merged = list(dict.fromkeys(list(self.favorites) + remote))
                self.favorites = merged
                changed += 1
            elif fname == 'pins.json' and isinstance(data, list):
                remote = [str(x) for x in data if isinstance(x, (str, int, float))]
                merged = list(dict.fromkeys(list(self.pins) + remote))
                self.pins = merged
                changed += 1
            elif fname == 'tags.json' and isinstance(data, dict):
                # merge tag sets per item
                for k, v in data.items():
                    if not isinstance(k, str):
                        continue
                    rv = [str(tg) for tg in (v or []) if isinstance(tg, (str, int, float))]
                    cur = self.tags.get(k, [])
                    merged = list(dict.fromkeys(cur + rv))
                    if merged:
                        self.tags[k] = merged
                changed += 1
            elif fname == 'expiry.json' and isinstance(data, dict):
                for k, v in data.items():
                    if not isinstance(k, str):
                        continue
                    try:
                        ts = float(v)
                    except Exception:
                        continue
                    if k not in self.expiry:
                        self.expiry[k] = ts
                    else:
                        # conservative: earliest expiry wins
                        self.expiry[k] = min(float(self.expiry.get(k) or ts), ts)
                changed += 1
            elif fname == 'formats.json' and isinstance(data, dict):
                # merge stored rich formats (prefer local if present; fill missing from remote)
                try:
                    for k, v in data.items():
                        if not isinstance(k, str):
                            continue
                        if not isinstance(v, dict):
                            continue
                        cur = (getattr(self, 'clip_formats', {}) or {}).get(k)
                        if not isinstance(cur, dict):
                            cur = {}
                        merged = dict(cur)
                        if 'html_b64' not in merged and isinstance(v.get('html_b64'), str):
                            merged['html_b64'] = v.get('html_b64')
                        if 'rtf_b64' not in merged and isinstance(v.get('rtf_b64'), str):
                            merged['rtf_b64'] = v.get('rtf_b64')
                        if merged:
                            (getattr(self, 'clip_formats', {}) or {})[k] = merged
                except Exception:
                    pass
                changed += 1

            last[fname] = r_mtime

        self._sync_seen_mtimes = last

        if changed:
            # apply expiries immediately and refresh UI
            self._purge_expired(silent=True)
            try:
                self._refresh_tag_filter_values()
            except Exception:
                pass
            self._refresh_list(select_last=False)
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass


        return changed

    def _sync_push(self, folder: Path) -> int:
        # Copy local files to remote if local newer
        pushed = 0
        for local_path, fname in self._sync_paths():
            src = local_path
            dst = folder / fname
            try:
                if not src.exists():
                    continue
                src_m = src.stat().st_mtime
                dst_m = dst.stat().st_mtime if dst.exists() else 0
                if src_m <= dst_m:
                    continue
                tmp = dst.with_suffix(dst.suffix + '.tmp')
                shutil.copy2(src, tmp)
                tmp.replace(dst)
                pushed += 1
            except Exception:
                continue
        return pushed

    # -----------------------------
    # Favorites & capacity management
    # -----------------------------
    def _ensure_favorites_present(self):
        """
        If a favorite exists but isn't in history anymore, we do not force-add it automatically
        (it may have been manually removed). This function is left as a hook if you want that behavior.
        """
        return

    def _prune_preserving_favorites(self, items: list[str], capacity: int) -> tuple[list[str], bool]:
        """
        Prune oldest items first, preserving Favorites and Pins.
        Returns: (new_items, success)
        """
        if len(items) <= capacity:
            return items, True

        # Protected items: Favorites, Pins, and anything with a Tag.
        keep = set(self.favorites) | set(self.pins)
        try:
            keep |= {k for k in (self.tags or {}).keys() if not str(k).startswith('IMG::')}
        except Exception:
            pass
        try:
            keep |= set((self.tags or {}).keys())
        except Exception:
            pass

        # Remove from the front (oldest) while over capacity, skipping Favorites/Pins
        i = 0
        out = items[:]
        while len(out) > capacity and i < len(out):
            if out[i] in keep:
                i += 1
                continue
            out.pop(i)

        if len(out) <= capacity:
            return out, True

        # If we are still over capacity, it means Favorites/Pins alone exceed capacity
        return out, False

    def _notify_limit_reached(self):
        if self._warned_limit_reached:
            return
        self._warned_limit_reached = True
        try:
            messagebox.showinfo(
                APP_NAME,
                f"You have reached your maximum stored items ({self.settings.max_history}).\n\n"
                "Consider increasing Max history in Settings, or Clean to remove non-favorites/non-pins.",
            )
        except Exception:
            pass

    def _notify_hard_cap_reached(self):
        try:
            messagebox.showwarning(
                APP_NAME,
                f"You have reached the hard-coded limit ({HARD_MAX_HISTORY}).\n\n"
                "You have used up all allocated memory for stored items.\n"
                "Please reconsider removing old copies from Favorites or reducing stored history.",
            )
        except Exception:
            pass

    def _notify_favorites_blocking(self):
        if self._warned_fav_block:
            return
        self._warned_fav_block = True
        try:
            messagebox.showwarning(
                APP_NAME,
                "Cannot add new clipboard item because your Favorites/Pins occupy the entire capacity.\n\n"
                "Increase Max history in Settings or remove some items from Favorites.",
            )
        except Exception:
            pass

    # -----------------------------
    # Clipboard polling
    # -----------------------------
    def _poll_clipboard(self):
        """Poll clipboard for new text (and optional images) and append to history."""
        try:
            # Expiry housekeeping (throttled)
            try:
                if time.time() - getattr(self, "_last_expiry_purge", 0.0) > 30:
                    self._purge_expired(silent=True)
                    self._last_expiry_purge = time.time()
            except Exception:
                pass

            if not self.paused:

                # Note: Lock is UI-only. Clipboard capture must continue while locked.

                # Optional: capture clipboard images
                try:
                    if self._adv('adv_images') and (bool(getattr(self, '_stores_loaded', True)) or not self._enc_all_enabled()):
                        img = self._clipboard_get_image()
                        if img is not None:
                            self._add_image_from_pil(img)
                except Exception:
                    pass

                payload = self._clipboard_get_rich_payload()
                text = payload.get('text') if isinstance(payload, dict) else ''
                text = text if isinstance(text, str) else ''

                if text and text != self.last_clip:
                    self.last_clip = text

                    # Store rich clipboard formats for this text (HTML/RTF) when present
                    try:
                        import base64
                        html = payload.get('html') if isinstance(payload, dict) else None
                        rtf = payload.get('rtf') if isinstance(payload, dict) else None
                        rec = {}
                        # Avoid huge blobs (keeps formats.json from exploding)
                        if isinstance(html, (bytes, bytearray)) and 1 <= len(html) <= 300_000:
                            rec['html_b64'] = base64.b64encode(bytes(html)).decode('ascii')
                        if isinstance(rtf, (bytes, bytearray)) and 1 <= len(rtf) <= 300_000:
                            rec['rtf_b64'] = base64.b64encode(bytes(rtf)).decode('ascii')
                        if rec:
                            (getattr(self, 'clip_formats', {}) or {})[text] = rec
                    except Exception:
                        pass

                    # If Encrypt-All stores are deferred (locked at startup), keep capturing but buffer in memory
                    try:
                        if not bool(getattr(self, '_stores_loaded', True)):
                            try:
                                self._captured_while_locked.append(text)
                            except Exception:
                                pass
                        else:
                            self._add_history_item(text)
                    except Exception:
                        self._add_history_item(text)
        except Exception:
            pass

        self._poll_job = self.after(self.settings.poll_ms, self._poll_clipboard)

    def _add_history_item(self, text: str):
        items = list(self.history)

        if items and items[-1] == text:
            return

        # Stable de-dupe: remove previous occurrences
        if text in items:
            items = [x for x in items if x != text]
        items.append(text)

        cap = self.settings.max_history

        # If at/over cap, prune preserving favorites
        if len(items) >= cap:
            self._notify_limit_reached()

        pruned, ok = self._prune_preserving_favorites(items, cap)

        if not ok:
            # Favorites/Pins exceed capacity; auto-expand capacity (up to hard cap) to avoid breaking capture.
            try:
                protected = set(self.favorites) | set(self.pins)
                try:
                    protected |= {k for k in (self.tags or {}).keys() if not str(k).startswith('IMG::')}
                except Exception:
                    pass
                need = len(protected) + 5
                if need > self.settings.max_history:
                    self.settings.max_history = min(HARD_MAX_HISTORY, max(self.settings.max_history, need))
                    # Try again with the expanded cap
                    cap = self.settings.max_history
                    pruned, ok = self._prune_preserving_favorites(items, cap)
                    if ok:
                        self.history = deque(pruned[-cap:], maxlen=cap)
                        self._refresh_list(select_last=True)
                        self._persist()
                        return
            except Exception:
                pass

            self._notify_favorites_blocking()
            if self.settings.max_history >= HARD_MAX_HISTORY:
                self._notify_hard_cap_reached()
            return

        items = pruned
        self.history = deque(items, maxlen=cap)
        self._refresh_list(select_last=True)
        self._persist()

    # -----------------------------
    # Expiry
    # -----------------------------
    def _purge_expired(self, silent: bool = True) -> int:
        """Remove items whose expiry timestamp has passed. Returns number removed."""
        try:
            now = time.time()
            expired = [k for k, ts in self.expiry.items() if isinstance(k, str) and isinstance(ts, (int, float)) and ts > 0 and ts <= now]
            if not expired:
                return 0
            exp_set = set(expired)

            # Remove from history
            items = [x for x in list(self.history) if x not in exp_set]
            self.history = deque(items, maxlen=self.settings.max_history)

            # Remove from favorites/pins
            self.favorites = [x for x in self.favorites if x not in exp_set]
            self.pins = [x for x in self.pins if x not in exp_set]

            # Remove tags/expiry
            for k in expired:
                self.tags.pop(k, None)
                self.expiry.pop(k, None)

            if not silent:
                messagebox.showinfo(APP_NAME, f"Removed {len(expired)} expired items.")

            self._refresh_list(select_last=False)
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            return len(expired)
        except Exception:
            return 0

    # -----------------------------
    # List/view helpers
    # -----------------------------
    def _format_list_item(self, item: str, display_index: int | None = None) -> str:
        """Format a history text item for the Listbox.

        display_index is the 1-based ordinal in the current view (used for line-number-like numbering).
        """
        one = re.sub(r"\s+", " ", item).strip()
        if len(one) > 90:
            one = one[:87] + "..."

        # Use emoji indicators (customizable via PIN_ICON/FAV_ICON/TAG_ICON)
        pin_mark = PIN_ICON if item in self.pins else " "
        fav_mark = FAV_ICON if item in self.favorites else " "
        tag_mark = TAG_ICON if self.tags.get(item) else " "

        idx = f"{display_index:>4} | " if isinstance(display_index, int) and display_index > 0 else ""
        prefix = f"{idx}{pin_mark}{fav_mark}{tag_mark} "
        return prefix + one

    def _format_list_item_image(self, key: str, rec: dict, display_index: int | None = None) -> str:
        """Format an image record for the Listbox (Images / mixed views)."""
        try:
            p = str(rec.get('path') or '')
            name = os.path.basename(p) if p else '(image)'
            ts = str(rec.get('created_at') or '')
            label = f"{IMAGE_ICON} {name}"
            if ts:
                label = f"{label}  —  {ts}"
        except Exception:
            label = f"{IMAGE_ICON} (image)"

        pin_mark = PIN_ICON if key in self.pins else ' ' 
        fav_mark = FAV_ICON if key in self.favorites else ' ' 
        tag_mark = TAG_ICON if self.tags.get(key) else ' ' 

        idx = f"{display_index:>4} | " if isinstance(display_index, int) and display_index > 0 else ''
        prefix = f"{idx}{pin_mark}{fav_mark}{tag_mark} "
        return prefix + label

    def _current_filter(self) -> str:
        if hasattr(self, "filter_var") and self.filter_var is not None:
            try:
                return str(self.filter_var.get())
            except Exception:
                return "all"
        return "all"

    def _get_all_tags(self):
        out = set()
        try:
            for _t, tags in (self.tags or {}).items():
                for tg in (tags or []):
                    if isinstance(tg, str) and tg.strip():
                        out.add(tg.strip())
        except Exception:
            pass
        return out

    def _refresh_tag_filter_values(self):
        # Refresh the Tag filter dropdown (if present)
        try:
            values = [""] + sorted(self._get_all_tags())
            if hasattr(self, "tag_combo") and self.tag_combo is not None:
                try:
                    self.tag_combo.configure(values=values)
                except Exception:
                    try:
                        self.tag_combo["values"] = values
                    except Exception:
                        pass
            if hasattr(self, "tag_filter_var") and self.tag_filter_var is not None:
                cur = str(self.tag_filter_var.get())
                if cur and cur not in values:
                    self.tag_filter_var.set("")
                    try:
                        self.filter_var.set("all")
                    except Exception:
                        pass
        except Exception:
            pass

    def _on_tag_filter_change(self):
        # When a tag is chosen, switch to Tag filter automatically
        try:
            tag = str(self.tag_filter_var.get()).strip()
        except Exception:
            tag = ""
        try:
            if tag:
                self.filter_var.set("tag")
            else:
                self.filter_var.set("all")
        except Exception:
            pass
        self._refresh_list(select_last=True)

    def _refresh_list(self, select_last: bool = False):
        """Refresh the left Listbox according to the current filter.

        Supports mixed views (text + images) for Favorites / Pins / Tags, and
        preserves selection across refreshes when possible.
        """
        # Keep the tag dropdown up-to-date
        try:
            self._refresh_tag_filter_values()
        except Exception:
            pass

        # Preserve current selection values (by item key/text)
        preserve = []
        try:
            for i in self.listbox.curselection():
                if 0 <= i < len(getattr(self, 'view_items', [])):
                    preserve.append(self.view_items[i])
        except Exception:
            preserve = []

        if not preserve:
            try:
                preserve = list(getattr(self, '_last_selected_texts', []) or [])
            except Exception:
                preserve = []
            if (not preserve) and getattr(self, '_selected_item_text', None):
                preserve = [self._selected_item_text]

        f = self._current_filter()

        # Build a stable global image-key map for mixed views
        self._image_map_all = {}
        try:
            for i, rec in enumerate(list(getattr(self, 'images', []) or [])):
                k = self._image_key_for_rec(rec, i)
                if k in self._image_map_all:
                    k = f"{k}::{i}"
                self._image_map_all[k] = rec
        except Exception:
            self._image_map_all = {}

        # Resolve base items for the filter
        items = []
        if f == 'img':
            self._image_map = dict(self._image_map_all)
            items = list(self._image_map.keys())
        elif f == 'fav':
            favs = list(getattr(self, 'favorites', []) or [])
            items = [x for x in favs if (x in self.history) or (x in self._image_map_all)]
        elif f == 'pin':
            pins = list(getattr(self, 'pins', []) or [])
            items = [x for x in pins if (x in self.history) or (x in self._image_map_all)]
        elif f == 'tag':
            try:
                tag = str(self.tag_filter_var.get()).strip()
            except Exception:
                tag = ''
            if tag:
                text_items = [x for x in list(self.history) if tag in self.tags.get(x, [])]
                img_items = [k for k in self._image_map_all.keys() if tag in self.tags.get(k, [])]
                items = text_items + img_items
            else:
                items = list(self.history)
        else:
            items = list(self.history)

        # Pinned items float to the top for all filters except the dedicated Pin view
        try:
            if f != 'pin':
                pinned = [x for x in (getattr(self, 'pins', []) or []) if x in items]
                pinned_set = set(pinned)
                rest = [x for x in items if x not in pinned_set]
                items = pinned + rest
        except Exception:
            pass

        self.view_items = items

        # Render
        self.listbox.delete(0, tk.END)
        for idx0, item in enumerate(self.view_items):
            if self._is_image_key(item):
                rec = self._image_map_all.get(item, {})
                self.listbox.insert(tk.END, self._format_list_item_image(item, rec, idx0 + 1))
            else:
                self.listbox.insert(tk.END, self._format_list_item(item, idx0 + 1))

            # Apply tag color (first matching tag with a configured color)
            try:
                for tg in self.tags.get(item, []):
                    col = getattr(self, 'tag_colors', {}).get(tg)
                    if isinstance(col, str) and col.strip():
                        self.listbox.itemconfig(idx0, fg=col.strip())
                        break
            except Exception:
                pass

        # Reset selection tracking
        self._prev_sel_set = set()
        self._sel_order = []

        # Restore selection
        if select_last and self.view_items:
            try:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(tk.END)
                self.listbox.see(tk.END)
            except Exception:
                pass
        elif preserve:
            try:
                self.listbox.selection_clear(0, tk.END)
                restored = []
                for t in preserve:
                    if t in self.view_items:
                        i = self.view_items.index(t)
                        self.listbox.selection_set(i)
                        restored.append(i)
                if restored:
                    self.listbox.see(restored[0])
                    self._prev_sel_set = set(restored)
                    self._sel_order = list(restored)
            except Exception:
                pass

        # Update preview for selection
        try:
            if self.listbox.curselection():
                self._on_select()
        except Exception:
            pass

        try:
            self._update_status_bar()
        except Exception:
            pass



    def _get_selected_indices(self) -> list[int]:
        return list(self.listbox.curselection())

    def _is_image_key(self, key) -> bool:
        try:
            if not (isinstance(key, str) and key.startswith('IMG::')):
                return False
            return key in (getattr(self, '_image_map_all', {}) or {})
        except Exception:
            return False

    def _get_selected_item(self):
        """Return a tuple: (kind, payload) where kind in {'text','image'}.

        - For text: payload is the text value (str)
        - For image: payload is the image record (dict) or None
        """
        sel = self._get_selected_indices()
        if not sel:
            return (None, None)
        i = sel[0]
        if i < 0 or i >= len(getattr(self, 'view_items', [])):
            return (None, None)
        key = self.view_items[i]
        if self._is_image_key(key):
            try:
                return ('image', getattr(self, '_image_map_all', {}).get(key))
            except Exception:
                return ('image', None)
        return ('text', key)

    def _get_selected_text(self) -> str | None:
        kind, payload = self._get_selected_item()
        if kind != 'text':
            return None
        return payload

    def _apply_reverse_lines(self, text: str) -> str:
        lines = text.splitlines()
        lines.reverse()
        return "\n".join(lines)

    def _get_preview_display_text_for_item(self, item_text: str) -> str:
        try:
            if getattr(self, '_reverse_item_text', None) == item_text:
                return self._apply_reverse_lines(item_text)
        except Exception:
            pass
        return item_text

    def _set_preview_text(self, text: str, mark_clean: bool = True):
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", text)
        self._preview_dirty = not mark_clean
        self._update_preview_dirty_ui()

        # Update line numbers gutter (if present)
        try:
            self._update_preview_line_numbers()
        except Exception:
            pass

        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _update_preview_line_numbers(self):
        """Render 1-based line numbers next to the Preview Text widget."""
        if not hasattr(self, 'preview_gutter') or self.preview_gutter is None:
            return
        try:
            # Determine number of lines (end-1c avoids trailing newline)
            end_index = self.preview.index('end-1c')
            lines = int(str(end_index).split('.')[0])
        except Exception:
            lines = 1

        nums = "\n".join(str(i) for i in range(1, max(1, lines) + 1)) + "\n"
        try:
            self.preview_gutter.configure(state='normal')
            self.preview_gutter.delete('1.0', tk.END)
            self.preview_gutter.insert('1.0', nums)
            self.preview_gutter.configure(state='disabled')
        except Exception:
            pass

        # Keep gutter scrolled in sync with preview
        try:
            first, _last = self.preview.yview()
            self.preview_gutter.yview_moveto(first)
        except Exception:
            pass

    def _show_text_preview(self):
        """Show the text preview frame (and hide the image preview frame)."""
        try:
            if hasattr(self, 'preview_image_frame') and self.preview_image_frame is not None:
                self.preview_image_frame.grid_remove()
        except Exception:
            pass
        try:
            if hasattr(self, 'preview_text_frame') and self.preview_text_frame is not None:
                self.preview_text_frame.grid()
        except Exception:
            pass

    def _show_image_preview(self, rec: dict | None):
        """Show an image record in the preview pane."""
        # Hide text editor UI
        try:
            if hasattr(self, 'preview_text_frame') and self.preview_text_frame is not None:
                self.preview_text_frame.grid_remove()
        except Exception:
            pass
        try:
            if hasattr(self, 'preview_image_frame') and self.preview_image_frame is not None:
                self.preview_image_frame.grid()
        except Exception:
            return

        if not rec:
            try:
                self.image_preview_title.configure(text='(No image selected)')
                self.image_preview_label.configure(image='', text='')
            except Exception:
                pass
            return

        path = str(rec.get('path') or '').strip()
        title = os.path.basename(path) if path else '(image)'
        try:
            self.image_preview_title.configure(text=title)
        except Exception:
            pass

        if not path or not os.path.exists(path):
            try:
                self.image_preview_label.configure(text='Image file not found.', image='')
            except Exception:
                pass
            return

        # Load & scale the image to the available preview area
        try:
            import io
            b = self._load_image_bytes(rec)
            if b is None:
                raise RuntimeError("no image bytes")
            img = Image.open(io.BytesIO(b))
            # Determine target size
            self.preview_container.update_idletasks()
            w = int(self.preview_container.winfo_width() or 800)
            h = int(self.preview_container.winfo_height() or 600)
            # Leave room for title + padding
            h = max(100, h - 60)
            w = max(100, w - 40)
            img.thumbnail((w, h))
            self._img_preview_tk = ImageTk.PhotoImage(img)
            self.image_preview_label.configure(image=self._img_preview_tk, text='')
        except Exception:
            try:
                self.image_preview_label.configure(text='Could not render image preview.', image='')
            except Exception:
                pass

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

        kind, payload = self._get_selected_item()
        if not kind:
            self._selected_item_text = None
            self._show_text_preview()
            self._set_preview_text("", mark_clean=True)
            return

        if kind == 'image':
            self._selected_item_text = None
            try:
                self.reverse_var.set(False)
            except Exception:
                pass
            self._show_image_preview(payload)
            return

        t = str(payload)
        self._selected_item_text = t

        # Keep reverse-lines checkbox in sync with the selected item
        try:
            self.reverse_var.set(getattr(self, '_reverse_item_text', None) == t)
        except Exception:
            pass
        display = self._get_preview_display_text_for_item(t)
        self._show_text_preview()
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

        # Cache selected texts so refreshes after button actions can restore selection
        try:
            self._last_selected_texts = [self.view_items[i] for i in sorted(current) if 0 <= i < len(self.view_items)]
        except Exception:
            self._last_selected_texts = []

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
        # Reverse-lines is per-selected-item (does not affect future clipboard items).
        if self._selected_item_text is None:
            try:
                self.reverse_var.set(False)
            except Exception:
                pass
            return

        want = bool(self.reverse_var.get())
        if want:
            # Enable reverse-lines for the currently selected history item.
            self._reverse_item_text = self._selected_item_text
        else:
            # Disable only if the current item is the active reversed one.
            if getattr(self, '_reverse_item_text', None) == self._selected_item_text:
                self._reverse_item_text = None

        display = self._get_preview_display_text_for_item(self._selected_item_text)
        if self._preview_dirty:
            resp = messagebox.askyesno(APP_NAME, "Reverse-lines will re-render Preview.\nDiscard current unsaved edits?")
            if not resp:
                # Restore checkbox to the current item's real state
                try:
                    self.reverse_var.set(getattr(self, '_reverse_item_text', None) == self._selected_item_text)
                except Exception:
                    pass
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

        # Keep line numbers up to date while editing
        try:
            self._update_preview_line_numbers()
        except Exception:
            pass

        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _save_preview_edits(self, _event=None):
        text_now = self.preview.get("1.0", tk.END).rstrip("\n")

        if self._selected_item_text is None:
            if text_now.strip():
                self._add_history_item(text_now)
                self.status_var.set(f"Saved new item from Preview — {now_ts()}")
                self._preview_dirty = False
                self._refresh_tag_filter_values()
            self._update_preview_dirty_ui()

            # Apply theme colors to tk widgets and initialize status bar
            try:
                self._apply_theme_to_tk_widgets()
            except Exception:
                pass
            try:
                self._update_status_bar()
            except Exception:
                pass
            return

        old = self._selected_item_text
        new = text_now

        if not new.strip():
            messagebox.showwarning(APP_NAME, "Cannot save an empty item.")
            return

        items = [x for x in self.history if x != old]
        items.append(new)

        # favorites map update
        if old in self.favorites:
            self.favorites = [new if x == old else x for x in self.favorites]

        # prune if needed (preserve favorites)
        cap = self.settings.max_history
        pruned, ok = self._prune_preserving_favorites(items, cap)
        if not ok:
            self._notify_favorites_blocking()
            return

        self.history = deque(pruned, maxlen=cap)
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
        # If an image is selected in the Images view, copy that image (or its path as a fallback).
        kind, payload = self._get_selected_item()
        if kind == 'image':
            ok = self._copy_selected_image(payload)
            if ok:
                try:
                    self.status_var.set(f"Copied image — {now_ts()}")
                except Exception:
                    pass
            else:
                messagebox.showerror(APP_NAME, "Failed to copy image to clipboard.")
            return

        # Text: if the user has selected a range in Preview, copy only that range.
        # Otherwise, copy the selected history item (with rich formatting if available).
        try:
            if self.preview.tag_ranges('sel'):
                out = self.preview.get('sel.first', 'sel.last')
            else:
                st = self._get_selected_text()
                if isinstance(st, str) and st:
                    if self._clipboard_set_rich_text(st):
                        try:
                            self.status_var.set(f"Copied — {now_ts()}")
                        except Exception:
                            pass
                        return
                out = self.preview.get("1.0", tk.END)
        except Exception:
            out = self.preview.get("1.0", tk.END)

        out = (out or '').rstrip("\n")
        if not out.strip():
            return
        if self._clipboard_set_text(out):
            try:
                self.status_var.set(f"Copied — {now_ts()}")
            except Exception:
                pass
        else:
            messagebox.showerror(APP_NAME, "Failed to copy to clipboard.")

    def _delete_selected(self):
        kind, payload = self._get_selected_item()
        if kind == 'image':
            rec = payload
            if not rec:
                return
            if not messagebox.askyesno(APP_NAME, "Delete selected image from the list?\n\n(Note: this will also attempt to delete the image file.)"):
                return
            try:
                path = str(rec.get('path') or '').strip()
            except Exception:
                path = ''

            key = self._image_key_for_rec(rec)
            try:
                if key in self.favorites:
                    self.favorites = [x for x in self.favorites if x != key]
                if key in self.pins:
                    self.pins = [x for x in self.pins if x != key]
                self.tags.pop(key, None)
                self.expiry.pop(key, None)
            except Exception:
                pass

            try:
                self.images = [r for r in (getattr(self, 'images', []) or []) if r is not rec and r.get('id') != rec.get('id')]
                self._save_images_meta()
            except Exception:
                pass
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass

            try:
                if hasattr(self, 'filter_var') and str(self.filter_var.get()) == 'img':
                    self._refresh_list(select_last=True)
            except Exception:
                pass
            return

        t = self._get_selected_text()
        if not t:
            return

        items = [x for x in self.history if x != t]
        self.history = deque(items, maxlen=self.settings.max_history)

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]

        if t in self.pins:
            self.pins = [x for x in self.pins if x != t]

        self.tags.pop(t, None)
        self.expiry.pop(t, None)

        if self._selected_item_text == t:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)

        self._refresh_list(select_last=True)
        self._persist()

    def _clean_keep_favorites(self):
        """Clean action.

        Requirement:
        - NOTHING that has a Favorite, Pin, or ANY Tag should be removed.
        - Images follow the same rule.
        """
        if not messagebox.askyesno(APP_NAME, "Clean history and keep Favorites/Pins/Tagged items?\n\nAnything with a tag, pin, or favorite will be preserved."):
            return

        # Protected keys: favorites, pins, and any key that has at least one tag.
        keep = set(self.favorites) | set(self.pins)
        try:
            keep |= {k for k, v in (self.tags or {}).items() if isinstance(v, list) and len(v) > 0}
        except Exception:
            pass

        # Text history: keep only protected items (favorites/pins/tagged)
        kept = [x for x in self.history if x in keep]
        self.history = deque(kept, maxlen=self.settings.max_history)

        # Images: keep only protected; remove files + metadata + tag/expiry refs for removed images
        try:
            new_images = []
            for i, rec in enumerate(list(getattr(self, 'images', []) or [])):
                k = self._image_key_for_rec(rec, i)
                if k in keep:
                    new_images.append(rec)
                else:
                    # Remove file on disk
                    try:
                        p = str(rec.get('path') or '').strip()
                        if p and os.path.exists(p):
                            os.remove(p)
                    except Exception:
                        pass
                    # Remove metadata pointers for removed item only
                    try:
                        self.tags.pop(k, None)
                        self.expiry.pop(k, None)
                    except Exception:
                        pass
                    try:
                        if k in self.favorites:
                            self.favorites = [x for x in self.favorites if x != k]
                        if k in self.pins:
                            self.pins = [x for x in self.pins if x != k]
                    except Exception:
                        pass

            self.images = new_images
            self._save_images_meta()
        except Exception:
            pass

        # Remove tag/expiry entries for text items that no longer exist in history
        try:
            alive_text = set(self.history)
            for k in list((self.tags or {}).keys()):
                if str(k).startswith('IMG::'):
                    continue
                if k not in alive_text and k not in keep:
                    self.tags.pop(k, None)
                    self.expiry.pop(k, None)
        except Exception:
            pass

        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._refresh_list(select_last=True)
        self._persist()
        self.status_var.set(f"Cleaned (kept favorites/pins/tagged) — {now_ts()}")

    def _toggle_favorite_selected(self):
        sel = self._get_selected_indices()
        if not sel:
            return
        items = [self.view_items[i] for i in sel if 0 <= i < len(self.view_items)]
        if not items:
            return

        any_unfav = any(it not in self.favorites for it in items)
        if any_unfav:
            for it in items:
                if it not in self.favorites:
                    self.favorites.append(it)
        else:
            remove_set = set(items)
            self.favorites = [x for x in self.favorites if x not in remove_set]

        self._refresh_list()
        self._persist()


    def _toggle_pin_selected(self):
        """Pin/Unpin selected items. Pins are preserved during pruning and float to the top."""
        sel = self._get_selected_indices()
        if not sel:
            return
        texts = [self.view_items[i] for i in sel if 0 <= i < len(self.view_items)]
        if not texts:
            return

        # If any item is not pinned -> pin all; else unpin all
        any_unpinned = any(t not in self.pins for t in texts)
        if any_unpinned:
            for t in texts:
                if t not in self.pins:
                    self.pins.append(t)
        else:
            self.pins = [x for x in self.pins if x not in set(texts)]

        self._refresh_list()
        self._persist()

    def _open_tags_dialog(self):
        sel = self._get_selected_indices()
        if not sel:
            return
        texts = [self.view_items[i] for i in sel if 0 <= i < len(self.view_items)]
        if not texts:
            return

        dlg = tk.Toplevel(self)
        dlg.title('Tags')
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry('420x340')

        tk.Label(dlg, text=f'Selected items: {len(texts)}', font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=12, pady=(12,6))

        # Current tags (intersection)
        current_sets = []
        for t in texts:
            current_sets.append(set(self.tags.get(t, [])))
        common = set.intersection(*current_sets) if current_sets else set()

        tk.Label(dlg, text='Common tags on selection:').pack(anchor='w', padx=12)
        common_var = tk.StringVar(value=', '.join(sorted(common)) if common else '(none)')
        tk.Label(dlg, textvariable=common_var).pack(anchor='w', padx=12, pady=(0,8))

        frm = tk.Frame(dlg)
        frm.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        frm.columnconfigure(0, weight=1)

        tk.Label(frm, text='Add tag:').grid(row=0, column=0, sticky='w')
        tag_var = tk.StringVar(value='')
        ent = tk.Entry(frm, textvariable=tag_var)
        ent.grid(row=1, column=0, sticky='ew', pady=(0,8))
        ent.focus_set()

        all_tags = sorted({tag for tags in self.tags.values() for tag in tags})
        tk.Label(frm, text='Existing tags:').grid(row=2, column=0, sticky='w')
        lb = tk.Listbox(frm, height=8, exportselection=False)
        lb.grid(row=3, column=0, sticky='nsew')
        frm.rowconfigure(3, weight=1)
        for tg in all_tags:
            lb.insert(tk.END, tg)

        def _apply_tag_color_styling():
            """Colorize tags in the tag list based on configured tag colors."""
            try:
                lb_end = lb.size()
                for i in range(lb_end):
                    tg = lb.get(i)
                    col = getattr(self, 'tag_colors', {}).get(tg)
                    if isinstance(col, str) and col.strip():
                        try:
                            lb.itemconfig(i, fg=col.strip())
                        except Exception:
                            pass
                    else:
                        try:
                            lb.itemconfig(i, fg='')
                        except Exception:
                            pass
            except Exception:
                pass

        _apply_tag_color_styling()

        btns = tk.Frame(dlg)
        btns.pack(fill=tk.X, padx=12, pady=(0,12))

        def add_tag():
            tg = tag_var.get().strip()
            if not tg:
                # allow picking from listbox
                try:
                    idx = lb.curselection()
                    if idx:
                        tg = lb.get(idx[0]).strip()
                except Exception:
                    tg = ''
            if not tg:
                return
            for t in texts:
                cur = self.tags.get(t, [])
                if tg not in cur:
                    cur = cur + [tg]
                    self.tags[t] = cur
            self._refresh_tag_filter_values()
            self._refresh_list()
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            # refresh displayed common tags
            sets = [set(self.tags.get(t, [])) for t in texts]
            com = set.intersection(*sets) if sets else set()
            common_var.set(', '.join(sorted(com)) if com else '(none)')
            if tg not in all_tags:
                lb.insert(tk.END, tg)

        def remove_tag():
            # remove selected existing tag from selected clips
            tg = ''
            try:
                idx = lb.curselection()
                if idx:
                    tg = lb.get(idx[0]).strip()
            except Exception:
                tg = ''
            if not tg:
                return
            for t in texts:
                cur = [x for x in self.tags.get(t, []) if x != tg]
                if cur:
                    self.tags[t] = cur
                else:
                    self.tags.pop(t, None)
            self._refresh_tag_filter_values()
            self._refresh_list()
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            sets = [set(self.tags.get(t, [])) for t in texts]
            com = set.intersection(*sets) if sets else set()
            common_var.set(', '.join(sorted(com)) if com else '(none)')


        def set_tag_color():
            tg = tag_var.get().strip()
            if not tg:
                try:
                    idx = lb.curselection()
                    if idx:
                        tg = lb.get(idx[0]).strip()
                except Exception:
                    tg = ''
            if not tg:
                return
            try:
                from tkinter import colorchooser
                picked = colorchooser.askcolor(title=f'Select color for tag: {tg}')
                if not picked or not picked[1]:
                    return
                self.tag_colors[tg] = picked[1]
                self._persist()
                _apply_tag_color_styling()
                self._refresh_list()
                try:
                    self._schedule_inactivity_lock()
                except Exception:
                    pass
            except Exception:
                pass

        def clear_tag_color():
            tg = tag_var.get().strip()
            if not tg:
                try:
                    idx = lb.curselection()
                    if idx:
                        tg = lb.get(idx[0]).strip()
                except Exception:
                    tg = ''
            if not tg:
                return
            try:
                self.tag_colors.pop(tg, None)
                self._persist()
                _apply_tag_color_styling()
                self._refresh_list()
                try:
                    self._schedule_inactivity_lock()
                except Exception:
                    pass
            except Exception:
                pass

        tk.Button(btns, text='Add/Apply', command=add_tag).pack(side=tk.LEFT)
        tk.Button(btns, text='Remove', command=remove_tag).pack(side=tk.LEFT, padx=(8,0))
        tk.Button(btns, text='Set Color…', command=set_tag_color).pack(side=tk.LEFT, padx=(8,0))
        tk.Button(btns, text='Clear Color', command=clear_tag_color).pack(side=tk.LEFT, padx=(8,0))
        tk.Button(btns, text='Close', command=dlg.destroy).pack(side=tk.RIGHT)

    def _set_expiry_selected(self):
        sel = self._get_selected_indices()
        if not sel:
            return
        texts = [self.view_items[i] for i in sel if 0 <= i < len(self.view_items)]
        if not texts:
            return

        mins = simpledialog.askinteger(APP_NAME, 'Expire selected items after how many minutes?\n\n(Example: 5, 60, 1440)', minvalue=1, maxvalue=525600)
        if not mins:
            return
        ts = time.time() + (int(mins) * 60)
        for t in texts:
            self.expiry[t] = ts

        self._persist()
        self.status_var.set(f'Expiry set ({mins} min) — {now_ts()}')

    def _clear_expiry_selected(self):
        sel = self._get_selected_indices()
        if not sel:
            return
        texts = [self.view_items[i] for i in sel if 0 <= i < len(self.view_items)]
        if not texts:
            return
        for t in texts:
            self.expiry.pop(t, None)
        self._persist()
        self.status_var.set(f'Expiry cleared — {now_ts()}')

    def _paste_last(self, do_type: bool = False):
        if not self.history:
            return
        t = list(self.history)[-1]
        try:
            self._clipboard_set_rich_text(t)
        except Exception:
            return
        if _kbd is None:
            return
        try:
            # Small delay lets focus return to last app
            def later():
                try:
                    if do_type:
                        _kbd.write(t)
                    else:
                        _kbd.send('ctrl+v')
                except Exception:
                    pass
            self.after(80, later)
        except Exception:
            pass

    def _open_quick_paste(self):
        """Keyboard-first quick paste/typing palette."""
        if not self._require_unlocked('Quick Paste'):
            return
        win = tk.Toplevel(self)
        win.title('Quick Paste')
        win.attributes('-topmost', True)
        win.geometry('620x420')

        q = tk.StringVar(value='')
        mode = tk.StringVar(value='paste')  # copy/paste/type

        top = tk.Frame(win)
        top.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(top, text='Search').pack(side=tk.LEFT)
        ent = tk.Entry(top, textvariable=q)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8,0))

        mid = tk.Frame(win)
        mid.pack(fill=tk.BOTH, expand=True, padx=10)
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)

        lb = tk.Listbox(mid, exportselection=False)
        # Readability + theme colors
        try:
            ui_font = ("Segoe UI", 11)
            lb.configure(font=ui_font)
            ent.configure(font=ui_font)
            colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
            if colors is not None:
                lb.configure(bg=colors.bg, fg=colors.fg, selectbackground=colors.primary, selectforeground=colors.light)
                win.configure(bg=colors.bg)
        except Exception:
            pass

        lb.grid(row=0, column=0, sticky='nsew')
        sb = tk.Scrollbar(mid, orient='vertical', command=lb.yview)
        sb.grid(row=0, column=1, sticky='ns')
        lb.configure(yscrollcommand=sb.set)

        bottom = tk.Frame(win)
        bottom.pack(fill=tk.X, padx=10, pady=10)
        tk.Radiobutton(bottom, text='Copy', value='copy', variable=mode).pack(side=tk.LEFT)
        tk.Radiobutton(bottom, text='Paste', value='paste', variable=mode).pack(side=tk.LEFT, padx=(10,0))
        tk.Radiobutton(bottom, text='Type', value='type', variable=mode).pack(side=tk.LEFT, padx=(10,0))
        tk.Button(bottom, text='Close', command=win.destroy).pack(side=tk.RIGHT)

        # Build list
        def get_items():
            items = list(self.history)
            pins = [x for x in items if x in self.pins]
            rest = [x for x in items if x not in self.pins]
            return pins + rest

        base = get_items()

        def refresh():
            term = q.get().strip().lower()
            lb.delete(0, tk.END)
            shown = 0
            for t in reversed(base):
                if term and term not in t.lower():
                    continue
                lb.insert(tk.END, self._format_list_item(t))
                shown += 1
                if shown >= 250:
                    break
            if shown:
                lb.selection_set(0)

        def selected_text():
            try:
                idx = lb.curselection()
                if not idx:
                    return None
                # map back to original item by stripping prefix match against base
                # simpler: rebuild displayed list mapping
                term = q.get().strip().lower()
                shown = []
                for t in reversed(base):
                    if term and term not in t.lower():
                        continue
                    shown.append(t)
                    if len(shown) >= 250:
                        break
                return shown[idx[0]] if 0 <= idx[0] < len(shown) else None
            except Exception:
                return None

        def commit():
            t = selected_text()
            if not t:
                return
            try:
                self._clipboard_set_rich_text(t)
            except Exception:
                return
            win.destroy()
            m = mode.get()
            if m == 'copy':
                return
            if _kbd is None:
                return
            def later():
                try:
                    if m == 'paste':
                        _kbd.send('ctrl+v')
                    elif m == 'type':
                        _kbd.write(t)
                except Exception:
                    pass
            try:
                self.after(90, later)
            except Exception:
                later()

        ent.bind('<KeyRelease>', lambda e: refresh())
        lb.bind('<Return>', lambda e: (commit(), 'break'))
        lb.bind('<Double-Button-1>', lambda e: (commit(), 'break'))
        ent.bind('<Return>', lambda e: (commit(), 'break'))
        win.bind('<Escape>', lambda e: win.destroy())

        refresh()
        ent.focus_set()

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

        self._selected_item_text = combined
        self._set_preview_text(self._get_preview_display_text_for_item(combined), mark_clean=True)
        self.status_var.set(f"Combined {len(parts)} items into new entry — {now_ts()}")

    # -----------------------------
    # Find / Highlight
    # -----------------------------

    def _fuzzy_match(self, q: str, text: str) -> bool:
        """Loose fuzzy match:
        - True if q is a substring of text (case-insensitive)
        - Or if the characters of q appear in-order within text
        """
        q = (q or "").strip().lower()
        if not q:
            return True
        t = (text or "").lower()
        if q in t:
            return True
        # subsequence match
        pos = 0
        for ch in q:
            pos = t.find(ch, pos)
            if pos == -1:
                return False
            pos += 1
        return True

    def _highlight_query_in_preview(self, query: str):
        """Highlight ALL matches of query in the preview (case-insensitive)."""
        try:
            self.preview.tag_remove("match", "1.0", tk.END)
        except Exception:
            return

        query = (query or "").strip()
        if not query:
            return

        text = self.preview.get("1.0", tk.END)
        if not text:
            return

        try:
            self.preview.tag_config("match", background="#2b78ff", foreground="white")
        except Exception:
            pass

        pattern = re.compile(re.escape(query), re.IGNORECASE)
        first = None
        for m in pattern.finditer(text):
            start_index = f"1.0+{m.start()}c"
            end_index = f"1.0+{m.end()}c"
            self.preview.tag_add("match", start_index, end_index)
            if first is None:
                first = start_index

        if first is not None:
            try:
                self.preview.see(first)
            except Exception:
                pass


    # -----------------------------
    # Format tools (Preview)
    # -----------------------------
    def _count_and_strip_invisible(self, text: str) -> tuple[str, int]:
        # Strips hidden/invisible characters commonly introduced by web/AI copy-paste.
        # - NBSP (\u00A0) is normalized to a normal space.
        # - Removes zero-width chars (\u200B,\u200C,\u200D), BOM (\uFEFF).
        # - Removes most ASCII control chars except for \n, \t, \r.
        # Returns (new_text, removed_count).
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        removed = 0
        out: list[str] = []
        for ch in text:
            o = ord(ch)
            if ch == "\u00A0":
                out.append(' ')
                removed += 1
                continue
            if ch in ("\u200B", "\u200C", "\u200D", "\uFEFF"):
                removed += 1
                continue
            if o < 32 and ch not in ('\n','\t','\r'):
                removed += 1
                continue
            out.append(ch)
        return ''.join(out), removed

    def _remove_blank_lines(self, text: str) -> tuple[str, int]:
        # Removes blank/whitespace-only lines. Returns (new_text, removed_lines).
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        lines = text.splitlines()
        kept: list[str] = []
        removed = 0
        for ln in lines:
            if ln.strip() == "":
                removed += 1
            else:
                kept.append(ln)
        new_text = '\n'.join(kept)
        if text.endswith('\n') and new_text and not new_text.endswith('\n'):
            new_text += '\n'
        return new_text, removed

    def _collapse_multiple_spaces(self, text: str) -> tuple[str, int]:
        """Collapse runs of 2+ spaces into a single space. Returns (new_text, removed_spaces)."""
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        # Only collapse regular spaces (not tabs). Keep line structure.
        new_text = re.sub(r" {2,}", " ", text)
        removed = max(0, len(text) - len(new_text))
        return new_text, removed

    def _trim_each_line(self, text: str) -> tuple[str, int]:
        """Trim leading/trailing whitespace on each line. Returns (new_text, removed_chars)."""
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        lines = text.splitlines(True)  # keep line endings
        out: list[str] = []
        removed = 0
        for ln in lines:
            # Separate newline to preserve exact endings
            if ln.endswith("\r\n"):
                body, nl = ln[:-2], "\r\n"
            elif ln.endswith("\n"):
                body, nl = ln[:-1], "\n"
            elif ln.endswith("\r"):
                body, nl = ln[:-1], "\r"
            else:
                body, nl = ln, ""
            trimmed = body.strip(" \t")
            removed += (len(body) - len(trimmed))
            out.append(trimmed + nl)
        return "".join(out), removed

    def _strip_trailing_whitespace(self, text: str) -> tuple[str, int]:
        """Remove trailing spaces/tabs at end of each line. Returns (new_text, removed_chars)."""
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        lines = text.splitlines(True)
        out: list[str] = []
        removed = 0
        for ln in lines:
            # Preserve newline sequence
            if ln.endswith("\r\n"):
                body, nl = ln[:-2], "\r\n"
            elif ln.endswith("\n"):
                body, nl = ln[:-1], "\n"
            elif ln.endswith("\r"):
                body, nl = ln[:-1], "\r"
            else:
                body, nl = ln, ""
            stripped = body.rstrip(" \t")
            removed += (len(body) - len(stripped))
            out.append(stripped + nl)
        return "".join(out), removed

    def _normalize_line_endings(self, text: str) -> tuple[str, int]:
        """Normalize CRLF/CR to LF. Returns (new_text, removed_cr_count)."""
        if not isinstance(text, str) or not text:
            return (text or ""), 0
        removed = text.count("\r")
        # Convert CRLF and CR to LF
        new_text = text.replace("\r\n", "\n").replace("\r", "\n")
        return new_text, removed

    def _apply_preview_text_change(self, new_text: str, status_note: str):
        """Apply a text change to the Preview editor and report via status bar."""
        try:
            self.preview.delete('1.0', tk.END)
            self.preview.insert('1.0', new_text)
            self._preview_dirty = True
            self._update_preview_dirty_ui()

            # Apply theme colors to tk widgets and initialize status bar
            try:
                self._apply_theme_to_tk_widgets()
            except Exception:
                pass
            try:
                self._update_status_bar()
            except Exception:
                pass
            if self.search_query:
                self._highlight_query_in_preview(self.search_query)
        except Exception:
            pass
        try:
            self._set_status_note(status_note)
        except Exception:
            try:
                self.status_var.set(status_note)
            except Exception:
                pass

    def _fmt_strip_hidden(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._count_and_strip_invisible(text)
        self._apply_preview_text_change(new_text, f"Format: removed {removed} hidden characters")

    def _fmt_remove_blank_lines(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._remove_blank_lines(text)
        self._apply_preview_text_change(new_text, f"Format: removed {removed} blank lines")

    def _fmt_strip_hidden_and_blanks(self):
        text = self.preview.get('1.0', tk.END)
        t1, removed_hidden = self._count_and_strip_invisible(text)
        t2, removed_blanks = self._remove_blank_lines(t1)
        self._apply_preview_text_change(t2, f"Format: removed {removed_blanks} blank lines, {removed_hidden} hidden characters")

    def _fmt_collapse_spaces(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._collapse_multiple_spaces(text)
        self._apply_preview_text_change(new_text, f"Format: collapsed spaces (removed {removed} spaces)")

    def _fmt_trim_each_line(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._trim_each_line(text)
        self._apply_preview_text_change(new_text, f"Format: trimmed lines (removed {removed} whitespace chars)")

    def _fmt_strip_trailing_ws(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._strip_trailing_whitespace(text)
        self._apply_preview_text_change(new_text, f"Format: removed {removed} trailing whitespace chars")

    def _fmt_normalize_line_endings(self):
        text = self.preview.get('1.0', tk.END)
        new_text, removed = self._normalize_line_endings(text)
        self._apply_preview_text_change(new_text, f"Format: normalized line endings (removed {removed} CR chars)")

    def _fmt_paste_plain_text(self):
        """Paste clipboard content into Preview after stripping hidden chars."""
        clip = self._clipboard_get_text()
        if not clip:
            self._set_status_note("Paste plain: clipboard empty")
            return
        clean, removed = self._count_and_strip_invisible(clip)
        try:
            # Replace selection if any
            try:
                start = self.preview.index("sel.first")
                end = self.preview.index("sel.last")
                self.preview.delete(start, end)
            except Exception:
                pass
            self.preview.insert(tk.INSERT, clean)
            self._preview_dirty = True
            self._update_preview_dirty_ui()
            if self.search_query:
                self._highlight_query_in_preview(self.search_query)
        except Exception:
            pass
        self._set_status_note(f"Paste plain: removed {removed} hidden characters")

    def _fmt_copy_preview_plain(self):
        """Copy a sanitized plain-text version of Preview to clipboard."""
        text = self.preview.get('1.0', tk.END)
        clean, removed = self._count_and_strip_invisible(text)
        ok = self._clipboard_set_text(clean)
        if ok:
            self._set_status_note(f"Copied plain text (removed {removed} hidden characters)")
        else:
            self._set_status_note("Copy plain failed")

    def _fmt_open_urls_in_preview(self):
        """Find http/https URLs in Preview and open them in the default browser."""
        text = self.preview.get('1.0', tk.END)
        # Basic URL matcher; trim common trailing punctuation
        raw = re.findall(r"https?://\S+", text)
        urls: list[str] = []
        for u in raw:
            u2 = u.rstrip(').,;\"\'<>]')
            if u2 not in urls:
                urls.append(u2)
        if not urls:
            self._set_status_note("Open URLs: none found")
            return
        # Prevent accidental flood
        max_open = 10
        to_open = urls[:max_open]
        if len(urls) > max_open:
            self._set_status_note(f"Open URLs: opening first {max_open} of {len(urls)}")
        else:
            self._set_status_note(f"Open URLs: opening {len(to_open)}")
        for u in to_open:
            try:
                webbrowser.open(u)
            except Exception:
                pass

    def _show_format_menu(self, anchor_widget):
        try:
            menu = getattr(self, '_format_menu', None)
            if menu is None:
                return
            x = anchor_widget.winfo_rootx()
            y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height()
            menu.tk_popup(x, y)
        except Exception:
            pass
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass


    def _search_live(self, _event=None):
        """Live search feedback while typing.

        Enhancements:
        - Searches text history AND image items
        - Matches against tags (both text and image keys)
        - Keeps results as item keys so jump works for images too
        """
        try:
            q = str(self.search_var.get()).strip()
        except Exception:
            q = ''
        self.search_query = q

        if not q:
            try:
                self.preview.tag_remove('match', '1.0', tk.END)
            except Exception:
                pass
            try:
                self._set_status_note('Search cleared')
            except Exception:
                pass
            self.search_matches = []
            self.search_index = 0
            return

        ql = q.lower()

        # Ensure image map exists (so images can be searched even when not in Images view)
        try:
            if not hasattr(self, '_image_map_all'):
                self._refresh_list()
        except Exception:
            pass

        matches = []

        # Text items
        for item in list(self.history):
            try:
                itl = item.lower()
                if ql in itl or self._fuzzy_match(q, item):
                    matches.append(item)
                    continue
                # Tag match
                for tg in self.tags.get(item, []) or []:
                    if isinstance(tg, str) and ql in tg.lower():
                        matches.append(item)
                        break
            except Exception:
                pass

        # Image keys
        try:
            for k, rec in (getattr(self, '_image_map_all', {}) or {}).items():
                try:
                    name = ''
                    if isinstance(rec, dict):
                        name = str(rec.get('id') or '') + ' ' + os.path.basename(str(rec.get('path') or ''))
                    if ql and name and ql in name.lower():
                        matches.append(k)
                        continue
                    for tg in self.tags.get(k, []) or []:
                        if isinstance(tg, str) and ql in tg.lower():
                            matches.append(k)
                            break
                except Exception:
                    pass
        except Exception:
            pass

        # De-dupe while preserving order
        seen=set()
        out=[]
        for x in matches:
            if x not in seen:
                out.append(x)
                seen.add(x)

        self.search_matches = out
        self.search_index = 0

        # Highlight in preview
        try:
            self._highlight_query_in_preview(q)
        except Exception:
            pass

        try:
            self._set_status_note(f"Search: {len(out)} match(es)")
        except Exception:
            pass

    def _search(self):
        q = self.search_var.get().strip()
        self.search_query = q
        self.search_matches = []
        self.search_index = 0

        if not q:
            self.status_var.set('Search cleared.')
            try:
                self.preview.tag_remove('match', '1.0', tk.END)
            except Exception:
                pass
            return

        ql = q.lower()

        # Ensure image map exists
        try:
            if not hasattr(self, '_image_map_all'):
                self._refresh_list()
        except Exception:
            pass

        matches = []

        # Text history matches
        for item in list(self.history):
            try:
                if ql in item.lower() or self._fuzzy_match(q, item):
                    matches.append(item)
                    continue
                for tg in self.tags.get(item, []) or []:
                    if isinstance(tg, str) and ql in tg.lower():
                        matches.append(item)
                        break
            except Exception:
                pass

        # Image matches (by id/filename or by tag)
        try:
            for k, rec in (getattr(self, '_image_map_all', {}) or {}).items():
                try:
                    name = ''
                    if isinstance(rec, dict):
                        name = str(rec.get('id') or '') + ' ' + os.path.basename(str(rec.get('path') or ''))
                    if name and ql in name.lower():
                        matches.append(k)
                        continue
                    for tg in self.tags.get(k, []) or []:
                        if isinstance(tg, str) and ql in tg.lower():
                            matches.append(k)
                            break
                except Exception:
                    pass
        except Exception:
            pass

        # De-dupe
        seen=set(); out=[]
        for x in matches:
            if x not in seen:
                out.append(x); seen.add(x)

        if not out:
            self.status_var.set(f'No matches for: {q}')
            try:
                self.preview.tag_remove('match', '1.0', tk.END)
            except Exception:
                pass
            return

        self.search_matches = out
        self.status_var.set(f'Found {len(out)} matches for: {q}')
        self._jump_to_item(out[0], highlight_query=q)

    def _jump_match(self, direction: int):
        if not self.search_matches:
            self._search()
            return
        self.search_index = (self.search_index + direction) % len(self.search_matches)
        self._jump_to_item(self.search_matches[self.search_index], highlight_query=self.search_query)

    def _jump_to_item(self, item_text: str, highlight_query: str = ''):
        # If this is an image key, ensure we're in a view that can show images
        try:
            if self._is_image_key(item_text):
                if self._current_filter() != 'img':
                    try:
                        self.filter_var.set('img')
                    except Exception:
                        pass
                    self._refresh_list()
        except Exception:
            pass

        if item_text not in self.view_items and self._current_filter() != 'all':
            try:
                # Fallback to all for text items
                if not self._is_image_key(item_text):
                    self.filter_var.set('all')
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

        # Preview: text shows text; images show thumbnail + metadata
        try:
            if self._is_image_key(item_text):
                kind, rec = self._get_selected_item()
                if kind == 'image' and isinstance(rec, dict):
                    self._show_image_in_preview(rec)
                else:
                    self._set_preview_text('(Image could not be loaded)', mark_clean=True)
            else:
                display = self._get_preview_display_text_for_item(item_text)
                self._set_preview_text(display, mark_clean=True)
        except Exception:
            try:
                display = self._get_preview_display_text_for_item(item_text)
                self._set_preview_text(display, mark_clean=True)
            except Exception:
                pass

        if highlight_query:
            self._highlight_query_in_preview(highlight_query)

    def _export(self):
        path = filedialog.asksaveasfilename(
            title="Export Data",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="copy2_export.json",
        )
        if not path:
            return

        payload = {
            "exported_at": now_ts(),
            "app_version": APP_VERSION,
            "history": list(self.history),
            "favorites": list(self.favorites),
            "pins": list(self.pins),
            "tags": dict(self.tags or {}),
            "tag_colors": dict(getattr(self, 'tag_colors', {}) or {}),
            "expiry": dict(getattr(self, 'expiry', {}) or {}),
            "settings": asdict(self.settings),
            "snippets": list(getattr(self, 'snippets', []) or []),
            "images": list(getattr(self, 'images', []) or []),
        }

        # Embed image bytes (best-effort) so exports are self-contained
        imgs_blob = []
        try:
            import base64
            for rec in list(getattr(self, 'images', []) or []):
                b = self._load_image_bytes(rec)
                if b is None:
                    continue
                imgs_blob.append({
                    'id': rec.get('id'),
                    'created_at': rec.get('created_at'),
                    'name': os.path.basename(str(rec.get('path') or '')),
                    'data_b64': base64.b64encode(b).decode('ascii'),
                })
        except Exception:
            imgs_blob = []
        payload['images_blob'] = imgs_blob

        out_obj = payload

        # Encrypt export if Encrypt-Exports OR Encrypt-All is enabled.
        try:
            want_enc = bool(getattr(self.settings, 'advanced_features', False) and (getattr(self.settings, 'adv_encrypt_exports', False) or getattr(self.settings, 'adv_encrypt_all_data', False)))
        except Exception:
            want_enc = False

        if want_enc:
            if Fernet is None:
                messagebox.showwarning(APP_NAME, 'Export encryption requires cryptography (Fernet). Exporting plaintext instead.')
            else:
                if not self._pin_is_set():
                    messagebox.showerror(APP_NAME, 'Export encryption requires a PIN. Set a PIN under Settings → Advanced.')
                    return
                pin = getattr(self, '_session_pin', None)
                if not pin:
                    pin = simpledialog.askstring(APP_NAME, 'Enter PIN to encrypt export:', show='*')
                    if not pin or not self._verify_pin_value(pin):
                        messagebox.showerror(APP_NAME, 'Incorrect PIN. Export cancelled.')
                        return
                    self._session_pin = str(pin).strip()
                out_obj = {
                    '__copy2_export_enc__': 1,
                    'env': self._encrypt_json_obj(payload, self._session_pin)
                }

        try:
            Path(path).write_text(json.dumps(out_obj, ensure_ascii=False, indent=2), encoding="utf-8")
            self.status_var.set(f"Exported — {path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Export failed:\n{e}")

    def _import(self):
        path = filedialog.askopenfilename(title="Import Data", filetypes=[("JSON", "*.json")])
        if not path:
            return

        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8", errors='ignore'))
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")
            return

        # Decrypt export envelope if needed
        data = raw
        try:
            if isinstance(raw, dict) and raw.get('__copy2_export_enc__') == 1 and isinstance(raw.get('env'), dict):
                if Fernet is None:
                    messagebox.showerror(APP_NAME, 'This export is encrypted but cryptography (Fernet) is not available.')
                    return
                if not self._pin_is_set():
                    messagebox.showerror(APP_NAME, 'This export is encrypted but no PIN is configured in this app instance.')
                    return
                pin = getattr(self, '_session_pin', None)
                if not pin:
                    pin = simpledialog.askstring(APP_NAME, 'Enter PIN to decrypt import:', show='*')
                    if not pin or not self._verify_pin_value(pin):
                        messagebox.showerror(APP_NAME, 'Incorrect PIN. Import cancelled.')
                        return
                    self._session_pin = str(pin).strip()
                data = self._decrypt_json_obj(raw['env'], self._session_pin)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")
            return

        try:
            history = data.get('history', []) if isinstance(data, dict) else []
            favorites = data.get('favorites', []) if isinstance(data, dict) else []
            pins = data.get('pins', []) if isinstance(data, dict) else []
            tags = data.get('tags', {}) if isinstance(data, dict) else {}
            tag_colors = data.get('tag_colors', {}) if isinstance(data, dict) else {}
            expiry = data.get('expiry', {}) if isinstance(data, dict) else {}
            snippets = data.get('snippets', []) if isinstance(data, dict) else []
            images_meta = data.get('images', []) if isinstance(data, dict) else []
            images_blob = data.get('images_blob', []) if isinstance(data, dict) else []

            # Merge history
            merged = list(self.history)
            if isinstance(history, list):
                for item in history:
                    if isinstance(item, str) and item.strip():
                        merged.append(item)

            # Stable de-dupe (keep latest occurrences)
            seen = set()
            out = []
            for item in reversed(merged):
                if item not in seen:
                    out.append(item)
                    seen.add(item)
            out.reverse()

            # Merge favorites/pins
            if isinstance(favorites, list):
                for x in favorites:
                    if isinstance(x, str) and x not in self.favorites:
                        self.favorites.append(x)
            if isinstance(pins, list):
                for x in pins:
                    if isinstance(x, str) and x not in self.pins:
                        self.pins.append(x)

            # Merge tags/colors/expiry
            if isinstance(tags, dict):
                for k, v in tags.items():
                    if isinstance(k, str) and isinstance(v, list):
                        cur = self.tags.get(k, []) if isinstance(self.tags, dict) else []
                        cur_set = set([t for t in cur if isinstance(t, str)])
                        for t in v:
                            if isinstance(t, str) and t not in cur_set:
                                cur.append(t)
                        self.tags[k] = cur
            if isinstance(tag_colors, dict):
                self.tag_colors.update({k: v for k, v in tag_colors.items() if isinstance(k, str) and isinstance(v, str)})
            if isinstance(expiry, dict):
                self.expiry.update({k: v for k, v in expiry.items() if isinstance(k, str)})

            # Merge snippets
            if isinstance(snippets, list):
                self.snippets = list(snippets)

            # Import images (prefer embedded blobs)
            try:
                import base64
                self.images_dir.mkdir(parents=True, exist_ok=True)
                existing_ids = set([str(r.get('id') or '') for r in (getattr(self, 'images', []) or [])])
                for blob in images_blob:
                    bid = str(blob.get('id') or '').strip()
                    if not bid:
                        continue
                    if bid in existing_ids:
                        continue
                    b = base64.b64decode(str(blob.get('data_b64') or ''))
                    fname = f"img_{bid}.png"
                    fpath = self.images_dir / fname
                    # Respect Encrypt-All
                    if self._enc_all_enabled() and getattr(self, '_session_pin', None) and Fernet is not None:
                        env = self._encrypt_json_obj({'data_b64': base64.b64encode(b).decode('ascii'), 'ext': 'png'}, self._session_pin)
                        fpath = self.images_dir / f"img_{bid}.c2img"
                        Path(fpath).write_text(json.dumps(env, ensure_ascii=False, indent=2), encoding='utf-8')
                    else:
                        Path(fpath).write_bytes(b)

                    rec = {'id': bid, 'path': str(fpath), 'created_at': blob.get('created_at') or now_ts()}
                    if not isinstance(getattr(self, 'images', None), list):
                        self.images = []
                    self.images.append(rec)
            except Exception:
                pass

            # Fall back: merge image metadata if present
            if isinstance(images_meta, list):
                # Only add those not already present
                existing = set([str(r.get('id') or '') for r in (getattr(self, 'images', []) or [])])
                for r in images_meta:
                    if not isinstance(r, dict):
                        continue
                    rid = str(r.get('id') or '').strip()
                    if rid and rid in existing:
                        continue
                    self.images.append(r)

            # Apply cap & prune preserving favorites/pins
            cap = self.settings.max_history
            pruned, ok = self._prune_preserving_favorites(out, cap)
            if not ok:
                self._notify_favorites_blocking()
            self.history = deque(pruned[-cap:], maxlen=cap)

            self._save_images_meta()
            self._refresh_list(select_last=True)
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            self.status_var.set(f"Imported — {path}")

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")
 
    def _open_settings(self, initial_tab: str = "General"):
        if not self._require_unlocked('Settings'):
            return
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry('760x560')
        dlg.minsize(720, 520)

        # Notebook tabs
        nb = ttk.Notebook(dlg)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tab_general = ttk.Frame(nb)
        tab_hotkeys = ttk.Frame(nb)
        tab_sync = ttk.Frame(nb)
        tab_advanced = ttk.Frame(nb)
        tab_help = ttk.Frame(nb)

        nb.add(tab_general, text='General')
        nb.add(tab_hotkeys, text='Hotkeys')
        nb.add(tab_sync, text='Sync')
        nb.add(tab_advanced, text='Advanced')
        nb.add(tab_help, text='Help')

        # -----------------
        # General tab
        # -----------------
        g = tab_general
        for c in range(3):
            g.columnconfigure(c, weight=(1 if c == 1 else 0))

        max_var = tk.StringVar(value=str(self.settings.max_history))
        poll_var = tk.StringVar(value=str(self.settings.poll_ms))
        sess_var = tk.BooleanVar(value=self.settings.session_only)
        upd_var = tk.BooleanVar(value=self.settings.check_updates_on_launch)

        ttk.Label(g, text=f"Max history (5–{HARD_MAX_HISTORY}):").grid(row=0, column=0, sticky='w', padx=10, pady=(12,6))
        ttk.Entry(g, textvariable=max_var, width=12).grid(row=0, column=1, sticky='w', padx=10, pady=(12,6))

        ttk.Label(g, text="Poll interval ms (100–5000):").grid(row=1, column=0, sticky='w', padx=10, pady=6)
        ttk.Entry(g, textvariable=poll_var, width=12).grid(row=1, column=1, sticky='w', padx=10, pady=6)

        ttk.Checkbutton(g, text="Session-only (do not save history)", variable=sess_var).grid(row=2, column=0, columnspan=2, sticky='w', padx=10, pady=(10,6))
        ttk.Checkbutton(g, text="Check updates on launch", variable=upd_var).grid(row=3, column=0, columnspan=2, sticky='w', padx=10, pady=6)

        theme_var = tk.StringVar(value=getattr(self.settings, 'theme', 'flatly'))
        themes = []
        if USE_TTKB and hasattr(self, 'style') and self.style is not None:
            try:
                themes = list(self.style.theme_names())
            except Exception:
                themes = []
        if not themes:
            themes = [getattr(self.settings, 'theme', 'flatly')]

        ttk.Button(g, text='Open Files Location', command=self._open_data_folder).grid(row=4, column=0, sticky='w', padx=10, pady=(10,6))
        ttk.Label(g, text='(Opens the AppData folder where Copy 2.0 stores its JSON / image files)').grid(row=5, column=0, columnspan=2, sticky='w', padx=10, pady=(0,8))

        ttk.Label(g, text="Theme:").grid(row=6, column=0, sticky='w', padx=10, pady=(10,6))
        cb_theme = ttk.Combobox(g, textvariable=theme_var, values=themes, width=22, state='readonly')
        cb_theme.grid(row=6, column=1, sticky='w', padx=10, pady=(10,6))
        if not USE_TTKB:
            ttk.Label(g, text="(Install ttkbootstrap to enable theme switching)").grid(row=7, column=0, columnspan=2, sticky='w', padx=10, pady=(6,0))

        # -----------------
        # Hotkeys tab
        # -----------------
        h = tab_hotkeys
        h.columnconfigure(1, weight=1)

        hk_en_var = tk.BooleanVar(value=getattr(self.settings, 'enable_global_hotkeys', False))
        hk_qp_var = tk.StringVar(value=getattr(self.settings, 'hotkey_quick_paste', 'ctrl+alt+v'))
        hk_pl_var = tk.StringVar(value=getattr(self.settings, 'hotkey_paste_last', 'ctrl+alt+shift+v'))

        ttk.Checkbutton(h, text="Enable global hotkeys (requires 'keyboard')", variable=hk_en_var).grid(row=0, column=0, columnspan=2, sticky='w', padx=10, pady=(12,8))
        ttk.Label(h, text="Quick Paste hotkey:").grid(row=1, column=0, sticky='w', padx=10, pady=6)
        ttk.Entry(h, textvariable=hk_qp_var, width=28).grid(row=1, column=1, sticky='w', padx=10, pady=6)
        ttk.Label(h, text="Paste last hotkey:").grid(row=2, column=0, sticky='w', padx=10, pady=6)
        ttk.Entry(h, textvariable=hk_pl_var, width=28).grid(row=2, column=1, sticky='w', padx=10, pady=6)

        # -----------------
        # Sync tab
        # -----------------
        s = tab_sync
        s.columnconfigure(1, weight=1)
        sync_en_var = tk.BooleanVar(value=getattr(self.settings, 'sync_enabled', False))
        sync_folder_var = tk.StringVar(value=getattr(self.settings, 'sync_folder', ''))
        sync_int_var = tk.StringVar(value=str(getattr(self.settings, 'sync_interval_sec', 10)))

        ttk.Checkbutton(s, text="Enable sync to folder", variable=sync_en_var).grid(row=0, column=0, columnspan=2, sticky='w', padx=10, pady=(12,8))
        ttk.Label(s, text="Sync folder:").grid(row=1, column=0, sticky='w', padx=10, pady=6)
        ttk.Entry(s, textvariable=sync_folder_var, width=46).grid(row=1, column=1, sticky='ew', padx=10, pady=6)

        def _browse_sync_folder():
            d = filedialog.askdirectory(title='Choose sync folder')
            if d:
                sync_folder_var.set(d)

        ttk.Button(s, text='Browse…', command=_browse_sync_folder).grid(row=2, column=1, sticky='w', padx=10, pady=(0,8))

        ttk.Label(s, text="Sync interval (sec):").grid(row=3, column=0, sticky='w', padx=10, pady=6)
        ttk.Entry(s, textvariable=sync_int_var, width=12).grid(row=3, column=1, sticky='w', padx=10, pady=6)

        # -----------------
        # Advanced tab
        # -----------------
        a = tab_advanced
        for c in range(3):
            a.columnconfigure(c, weight=(1 if c == 1 else 0))

        adv_master_var = tk.BooleanVar(value=getattr(self.settings, 'advanced_features', False))
        adv_lock_var = tk.BooleanVar(value=getattr(self.settings, 'adv_app_lock', False))
        adv_boot_var = tk.BooleanVar(value=getattr(self.settings, 'adv_start_on_boot', False))
        adv_encrypt_var = tk.BooleanVar(value=getattr(self.settings, 'adv_encrypt_exports', False))
        adv_all_var = tk.BooleanVar(value=getattr(self.settings, 'adv_encrypt_all_data', False))
        adv_images_var = tk.BooleanVar(value=getattr(self.settings, 'adv_images', False))
        adv_ss_var = tk.BooleanVar(value=getattr(self.settings, 'adv_screenshots', False))
        adv_snip_var = tk.BooleanVar(value=getattr(self.settings, 'adv_snippets', False))
        adv_trig_var = tk.BooleanVar(value=getattr(self.settings, 'adv_tmplt_trigger', False))
        trig_word_var = tk.StringVar(value=getattr(self.settings, 'tmplt_trigger_word', 'tmplt'))

        master_cb = ttk.Checkbutton(a, text='Enable Advanced Features (required for any options below)', variable=adv_master_var)
        master_cb.grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(12, 8))

        row = 1
        cb_lock = ttk.Checkbutton(a, text='App-level lock (PIN required on startup)', variable=adv_lock_var)
        cb_lock.grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=6)
        btn_pin = ttk.Button(a, text='Set / Change PIN…', command=lambda: self._set_or_change_pin_flow(parent=dlg))
        btn_pin.grid(row=row, column=2, sticky='e', padx=10, pady=6)

        # Separate Nuke PIN control (emergency wipe)
        row += 1
        ttk.Label(a, text='Nuke PIN (wipes all local + sync data)').grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=2)
        btn_nuke = ttk.Button(a, text='Set / Change NUKE PIN…', command=lambda: self._set_or_change_nuke_pin_flow(parent=dlg))
        btn_nuke.grid(row=row, column=2, sticky='e', padx=10, pady=2)
        row += 1

        # Inactivity auto-lock (UI-only; engine keeps running)
        cur_min = 0
        try:
            cur_min = int(getattr(self.settings, 'lock_timeout_minutes', 0) or 0)
        except Exception:
            cur_min = 0

        lock_timeout_mode_var = tk.StringVar()
        if cur_min <= 0:
            lock_timeout_mode_var.set('Never')
        elif cur_min == 5:
            lock_timeout_mode_var.set('5 min')
        else:
            lock_timeout_mode_var.set('Custom')

        lock_custom_min_var = tk.StringVar(value=str(cur_min if cur_min not in (0, 5) else 10))

        ttk.Label(a, text='Auto-lock after inactivity:').grid(row=row, column=0, sticky='w', padx=10, pady=6)
        cmb_lock = ttk.Combobox(a, textvariable=lock_timeout_mode_var, values=['Never', '5 min', 'Custom'], state='readonly', width=10)
        cmb_lock.grid(row=row, column=1, sticky='w', padx=10, pady=6)
        ent_lock_custom = ttk.Entry(a, textvariable=lock_custom_min_var, width=8)
        ent_lock_custom.grid(row=row, column=2, sticky='e', padx=10, pady=6)

        def _lock_timeout_ui_refresh(*_):
            try:
                if str(lock_timeout_mode_var.get()) == 'Custom':
                    ent_lock_custom.configure(state='normal')
                else:
                    ent_lock_custom.configure(state='disabled')
            except Exception:
                pass

        try:
            cmb_lock.bind('<<ComboboxSelected>>', _lock_timeout_ui_refresh)
        except Exception:
            pass
        _lock_timeout_ui_refresh()
        row += 1

        cb_boot = ttk.Checkbutton(a, text='Start on boot (Windows)', variable=adv_boot_var)
        cb_boot.grid(row=row, column=0, columnspan=3, sticky='w', padx=10, pady=6)
        row += 1

        cb_encrypt = ttk.Checkbutton(a, text='Encrypt exports (requires PIN / passphrase)', variable=adv_encrypt_var)
        cb_encrypt.grid(row=row, column=0, columnspan=3, sticky='w', padx=10, pady=6)
        row += 1

        cb_all = ttk.Checkbutton(a, text='Encrypt ALL local saved data (history/pins/tags/images/snippets) — requires PIN', variable=adv_all_var)
        cb_all.grid(row=row, column=0, columnspan=3, sticky='w', padx=10, pady=6)
        row += 1

        cb_images = ttk.Checkbutton(a, text='Capture images copied to clipboard', variable=adv_images_var)
        cb_images.grid(row=row, column=0, columnspan=3, sticky='w', padx=10, pady=6)
        row += 1

        cb_ss = ttk.Checkbutton(a, text='Enable screenshots (store as images)', variable=adv_ss_var)
        cb_ss.grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=6)
        btn_ss = ttk.Button(a, text='Capture Screenshot', command=self._capture_screenshot)
        btn_ss.grid(row=row, column=2, sticky='e', padx=10, pady=6)
        row += 1

        cb_snip = ttk.Checkbutton(a, text='Snippets / templates', variable=adv_snip_var)
        cb_snip.grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=6)
        btn_sn = ttk.Button(a, text='Manage Templates…', command=self._open_snippets_manager)
        btn_sn.grid(row=row, column=2, sticky='e', padx=10, pady=6)
        row += 1

        cb_trig = ttk.Checkbutton(a, text='Enable template trigger while typing (shows picker)', variable=adv_trig_var)
        cb_trig.grid(row=row, column=0, sticky='w', padx=10, pady=6)
        lbl_trig = ttk.Label(a, text='Trigger word:')
        lbl_trig.grid(row=row, column=1, sticky='e', padx=10, pady=6)
        ent_trig = ttk.Entry(a, textvariable=trig_word_var, width=14)
        ent_trig.grid(row=row, column=2, sticky='e', padx=10, pady=6)
        row += 1

        btn_open_img = ttk.Button(a, text='Open Images Folder', command=self._open_images_folder)
        btn_open_img.grid(row=row, column=2, sticky='e', padx=10, pady=(18, 0))
        lbl_img = ttk.Label(a, text='(Images and screenshots are stored here)')
        lbl_img.grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=(18, 0))

        def _adv_controls_state(*_):
            en = bool(adv_master_var.get())
            st = 'normal' if en else 'disabled'
            for w in a.winfo_children():
                if w is master_cb:
                    continue
                try:
                    w.configure(state=st)
                except Exception:
                    pass

        adv_master_var.trace_add('write', _adv_controls_state)
        _adv_controls_state()


        # -----------------
        # Help tab
        # -----------------
        hp = tab_help
        hp.columnconfigure(0, weight=1)
        hp.rowconfigure(0, weight=1)

        help_box = ttk.Frame(hp)
        help_box.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        help_box.columnconfigure(0, weight=1)
        help_box.rowconfigure(0, weight=1)

        ysb = ttk.Scrollbar(help_box, orient='vertical')
        ysb.grid(row=0, column=1, sticky='ns')
        txt = tk.Text(help_box, wrap='word', yscrollcommand=ysb.set, font=('Segoe UI', 10), padx=10, pady=10)
        txt.grid(row=0, column=0, sticky='nsew')
        ysb.config(command=txt.yview)
        # Accept either COPY2_HELP_TEXT (recommended) or a user-edited help_text block
        try:
            src = globals().get('help_text', None)
            if src is None:
                src = globals().get('COPY2_HELP_TEXT', '')
            help_str = _normalize_text_block(src).strip()
            if not help_str:
                help_str = '(Help content is empty.)'
            txt.insert('1.0', help_str + '\n')
        except Exception:
            txt.insert('1.0', '(Help content could not be loaded.)\n')
        txt.configure(state='disabled')

        # -----------------
        # Bottom buttons
        # -----------------
        bottom = ttk.Frame(dlg)
        bottom.pack(fill=tk.X, padx=10, pady=(0,10))
        bottom.columnconfigure(0, weight=1)

        def _apply_theme_now(new_theme: str):
            if not USE_TTKB:
                return
            st = getattr(self, 'style', None)
            if st is None:
                return
            try:
                st.theme_use(new_theme)
                self._active_theme = new_theme
                try:
                    self._apply_theme_to_tk_widgets()
                except Exception:
                    pass
                try:
                    self._set_status_note(f"Theme applied: {new_theme}")
                except Exception:
                    pass
            except Exception:
                pass

        def apply_settings(final: bool = False):
            """Apply settings immediately (autosave). If final=True, apply even if some fields are invalid."""
            # Numeric fields: only commit when valid to avoid noisy popups while typing
            try:
                mh = int(str(max_var.get()).strip())
                pm = int(str(poll_var.get()).strip())
            except Exception:
                if final:
                    return
                mh = None
                pm = None

            if mh is not None:
                mh = max(5, min(HARD_MAX_HISTORY, mh))
                if mh >= HARD_MAX_HISTORY:
                    try:
                        self._notify_hard_cap_reached()
                    except Exception:
                        pass
                self.settings.max_history = mh

            if pm is not None:
                pm = max(100, min(5000, pm))
                self.settings.poll_ms = pm

            self.settings.session_only = bool(sess_var.get())
            self.settings.check_updates_on_launch = bool(upd_var.get())

            # Theme
            if USE_TTKB:
                new_theme = str(theme_var.get()).strip() or 'flatly'
                if getattr(self.settings, 'theme', '') != new_theme:
                    self.settings.theme = new_theme
                    _apply_theme_now(new_theme)

            # Hotkeys
            self.settings.enable_global_hotkeys = bool(hk_en_var.get())
            self.settings.hotkey_quick_paste = str(hk_qp_var.get()).strip() or self.settings.hotkey_quick_paste
            self.settings.hotkey_paste_last = str(hk_pl_var.get()).strip() or self.settings.hotkey_paste_last

            # Sync
            self.settings.sync_enabled = bool(sync_en_var.get())
            self.settings.sync_folder = str(sync_folder_var.get()).strip()
            try:
                self.settings.sync_interval_sec = max(3, min(300, int(str(sync_int_var.get()).strip())))
            except Exception:
                pass

            # Advanced
            prev_master = bool(getattr(self.settings, 'advanced_features', False))
            prev_encrypt_all = bool(getattr(self.settings, 'adv_encrypt_all_data', False))

            self.settings.advanced_features = bool(adv_master_var.get())
            if self.settings.advanced_features:
                self.settings.adv_app_lock = bool(adv_lock_var.get())
                # Inactivity lock (UI-only). Only meaningful if App Lock is enabled.
                mins = 0
                try:
                    mode = str(lock_timeout_mode_var.get() or 'Never')
                    if mode == '5 min':
                        mins = 5
                    elif mode == 'Custom':
                        mins = int(str(lock_custom_min_var.get()).strip())
                        mins = max(1, min(24*60, mins))
                except Exception:
                    mins = 0
                self.settings.lock_timeout_minutes = mins if self.settings.adv_app_lock else 0
                self.settings.adv_start_on_boot = bool(adv_boot_var.get())
                self.settings.adv_encrypt_exports = bool(adv_encrypt_var.get())
                self.settings.adv_encrypt_all_data = bool(adv_all_var.get())
                self.settings.adv_images = bool(adv_images_var.get())
                self.settings.adv_screenshots = bool(adv_ss_var.get())
                self.settings.adv_snippets = bool(adv_snip_var.get())
                self.settings.adv_tmplt_trigger = bool(adv_trig_var.get())
                tw = str(trig_word_var.get() or '').strip()
                self.settings.tmplt_trigger_word = tw if tw else 'tmplt'
            else:
                # Turning off master disables all subordinate features
                self.settings.adv_app_lock = False
                self.settings.adv_start_on_boot = False
                self.settings.adv_encrypt_exports = False
                self.settings.adv_encrypt_all_data = False
                self.settings.lock_timeout_minutes = 0
                self.settings.adv_images = False
                self.settings.adv_screenshots = False
                self.settings.adv_snippets = False
                self.settings.adv_tmplt_trigger = False

            # Enforce PIN requirement for lock/encryption
            try:
                if self.settings.advanced_features and (self.settings.adv_app_lock or self.settings.adv_encrypt_all_data):
                    if not self._pin_is_set():
                        if messagebox.askyesno(APP_NAME, "This feature requires a PIN.\n\nSet a PIN now?"):
                            self._set_or_change_pin_flow(parent=dlg)
                    if not self._pin_is_set():
                        # Disable dependent flags
                        self.settings.adv_app_lock = False
                        self.settings.adv_encrypt_all_data = False
                        self.settings.lock_timeout_minutes = 0
                        adv_lock_var.set(False)
                        adv_all_var.set(False)
                        messagebox.showwarning(APP_NAME, 'Feature disabled because no PIN is set.')
            except Exception:
                pass

            # Start on boot (apply immediately)
            try:
                self._set_start_on_boot(bool(self.settings.advanced_features and self.settings.adv_start_on_boot))
            except Exception:
                pass

            # If encrypt-all changed, ensure we have session pin (prompt) and migrate stores best-effort.
            try:
                if prev_encrypt_all != bool(self.settings.advanced_features and self.settings.adv_encrypt_all_data):
                    if Fernet is None:
                        self.settings.adv_encrypt_all_data = False
                        adv_all_var.set(False)
                        messagebox.showwarning(APP_NAME, 'Encrypt-All requires cryptography (Fernet).')
                    else:
                        if not getattr(self, '_session_pin', None):
                            # Ask for PIN once
                            p = simpledialog.askstring(APP_NAME, 'Enter PIN to apply encryption changes:', show='*', parent=dlg)
                            if not p or not self._verify_pin_value(p):
                                messagebox.showerror(APP_NAME, 'Incorrect PIN. Encryption change was not applied.')
                                self.settings.adv_encrypt_all_data = prev_encrypt_all
                                adv_all_var.set(prev_encrypt_all)
                            else:
                                self._session_pin = str(p).strip()
                        # Migrate image files and re-save stores using the new encryption setting
                        try:
                            self._migrate_image_files_for_encrypt_all(bool(self.settings.advanced_features and self.settings.adv_encrypt_all_data))
                        except Exception:
                            pass
                        self._persist()
            except Exception:
                pass

            # Update inactivity timer immediately (if enabled)
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass

            # Re-bootstrap advanced hooks
            try:
                self._advanced_bootstrap()
            except Exception:
                pass

            # Apply history capacity (preserve favorites/pins)
            try:
                if mh is not None:
                    items = list(self.history)
                    pruned, ok = self._prune_preserving_favorites(items, mh)
                    if not ok:
                        self._notify_favorites_blocking()
                    self.history = deque(pruned[-mh:], maxlen=mh)
            except Exception:
                pass

            try:
                self._refresh_list(select_last=True)
            except Exception:
                pass

            # Persist + runtime features
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            try:
                self._register_global_hotkeys()
            except Exception:
                pass
            try:
                self._start_sync_job()
            except Exception:
                pass

        # Debounced autosave
        _autosave_state = {'job': None, 'guard': False}

        def schedule_autosave(*_):
            if _autosave_state['guard']:
                return
            try:
                if _autosave_state['job'] is not None:
                    dlg.after_cancel(_autosave_state['job'])
            except Exception:
                pass
            _autosave_state['job'] = dlg.after(450, lambda: apply_settings(final=False))

        # Wire traces
        for v in [max_var, poll_var, sess_var, upd_var, theme_var,
                  hk_en_var, hk_qp_var, hk_pl_var,
                  sync_en_var, sync_folder_var, sync_int_var,
                  adv_master_var, adv_lock_var, lock_timeout_mode_var, lock_custom_min_var, adv_boot_var, adv_encrypt_var, adv_all_var,
                  adv_images_var, adv_ss_var, adv_snip_var, adv_trig_var, trig_word_var]:
            try:
                v.trace_add('write', schedule_autosave)
            except Exception:
                pass

        # Apply once immediately to normalize and ensure defaults
        schedule_autosave()

        ttk.Button(bottom, text='Close', command=lambda: (apply_settings(final=True), dlg.destroy())).pack(side=tk.RIGHT)

        # Select initial tab
        tab_map = {'general': 0, 'hotkeys': 1, 'sync': 2, 'advanced': 3, 'help': 4}
        idx = tab_map.get(str(initial_tab or '').strip().lower(), 0)
        try:
            nb.select(idx)
        except Exception:
            pass

    # -----------------------------
    # Context menu
    # -----------------------------
    def _build_context_menu(self):
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copy", command=self._copy_selected)
        self.menu.add_command(label="Favorite / Unfavorite", command=self._toggle_favorite_selected)
        self.menu.add_command(label="Pin / Unpin", command=self._toggle_pin_selected)
        self.menu.add_command(label="Tags...", command=self._open_tags_dialog)
        self.menu.add_command(label="Set Expiry...", command=self._set_expiry_selected)
        self.menu.add_command(label="Clear Expiry", command=self._clear_expiry_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Save Preview Edit", command=self._save_preview_edits)
        self.menu.add_command(label="Revert Preview Edit", command=self._revert_preview_edits)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self._delete_selected)
        self.menu.add_command(label="Combine Selected", command=self._combine_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Quick Paste...", command=self._open_quick_paste)
        self.menu.add_separator()
        self.menu.add_command(label="Templates / Snippets...", command=self._open_snippets_manager)
        self.menu.add_command(label="Capture Screenshot", command=self._capture_screenshot)
        self.menu.add_command(label="Open Images Folder", command=self._open_images_folder)

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
    # Update system
    # -----------------------------
    def _check_updates_async(self, prompt_if_new: bool = True):
        def worker():
            try:
                self._log_check(f"Check start (installed=v{APP_VERSION})")
                info = get_latest_release_info()
                if not info:
                    self._log_check("Check failed: no response")
                    self.after(0, lambda: messagebox.showwarning(APP_NAME, "Could not check for updates."))
                    self.status_var.set(f"Update check failed — {now_ts()}")
                    return

                latest = str(info.get("version") or "").strip()
                html_url = str(info.get("html_url") or GITHUB_RELEASES_PAGE).strip()
                asset_url = str(info.get("asset_url") or "").strip()
                asset_name = str(info.get("asset_name") or "").strip()

                self._log_check(f"Latest={latest} asset={asset_name} url={asset_url}")

                if not latest or not is_newer_version(latest, f"v{APP_VERSION}"):
                    self.status_var.set(f"No updates available — {now_ts()}")
                    return

                if not prompt_if_new:
                    self.status_var.set(f"Update available: {latest} — {now_ts()}")
                    return

                def ui_prompt():
                    msg = (
                        "Update available.\n\n"
                        f"Installed: {APP_VERSION}\n"
                        f"Latest: {latest}\n\n"
                    )

                    if not asset_url:
                        msg += "No ZIP/EXE asset URL was found on the latest release.\nOpen the release page?"
                        yes = messagebox.askyesno(APP_NAME, msg)
                        if yes:
                            webbrowser.open(html_url)
                        return

                    msg += "Do you want to download and install the update now?\n\nYes = Auto-update\nNo = Cancel"
                    yes = messagebox.askyesno(APP_NAME, msg)
                    if yes:
                        self._download_and_apply_update_async(
                            {"version": latest, "asset_url": asset_url, "asset_name": asset_name, "html_url": html_url}
                        )
                    else:
                        self.status_var.set(f"Update skipped ({latest}) — {now_ts()}")

                self.after(0, ui_prompt)

            except Exception as e:
                self._log_check(f"ERROR: {repr(e)}")
                self.after(0, lambda: messagebox.showwarning(APP_NAME, "Could not check for updates."))
                self.status_var.set(f"Update check failed — {now_ts()}")

        threading.Thread(target=worker, daemon=True).start()

    def _download_and_apply_update_async(self, info: dict):
        """
        Robust updater:
        - Only auto-updates if frozen AND onefile build.
        - Downloads asset to a temp workdir
        - Stages Copy2.exe from zip (or direct exe)
        - Writes a BAT that:
            wait -> backup -> copy new -> env reset -> selftest retries -> rollback if needed
        - Starts BAT, then closes this instance.
        """
        def worker():
            asset_url = str(info.get("asset_url", "")).strip()
            latest = str(info.get("version", "")).strip()
            html_url = str(info.get("html_url", "")).strip()
            asset_name = str(info.get("asset_name", "")).strip()

            try:
                self._log_install(f"Update start -> latest={latest} url={asset_url}")

                if not is_frozen():
                    messagebox.showinfo(
                        APP_NAME,
                        "Auto-update is supported for the packaged portable .exe.\n\n"
                        "You are running a Python script environment; opening the release page instead.",
                    )
                    if html_url:
                        webbrowser.open(html_url)
                    return

                target_exe = Path(sys.executable)
                target_dir = target_exe.parent

                # Block onedir updates when release only ships EXE/ZIP without full folder payload
                if is_onedir_frozen(target_exe):
                    messagebox.showinfo(
                        APP_NAME,
                        "You appear to be running an ONEDIR build (folder-based).\n\n"
                        "Auto-update requires a ONEFILE build, or you must publish a full folder payload.\n"
                        "Opening the release page instead.",
                    )
                    if html_url:
                        webbrowser.open(html_url)
                    return

                if not target_exe.exists():
                    messagebox.showerror(APP_NAME, "Could not locate the current executable for updating.")
                    return

                self._log_install(f"Target exe: {target_exe}")

                workdir = Path(tempfile.mkdtemp(prefix="copy2_update_"))
                self._log_install(f"Workdir: {workdir}")

                download_path = workdir / (asset_name or "Copy2_update.bin")

                downloaded, expected = _http_download(asset_url, download_path, timeout=60)
                self._log_install(f"Downloaded bytes: {downloaded} expected={expected}")

                if expected is not None and downloaded != expected:
                    raise RuntimeError("Download size mismatch (possible partial download).")

                # Stage new exe
                staged_copy2 = workdir / "staged_Copy2.exe"
                staged_uninst = workdir / "staged_Copy2_Uninstall.exe"  # optional (may not exist)

                if str(download_path).lower().endswith(".exe"):
                    # direct exe
                    staged_copy2.write_bytes(download_path.read_bytes())
                    self._log_install(f"Staged exe direct: {staged_copy2} size={staged_copy2.stat().st_size}")

                else:
                    # zip path
                    if not zipfile.is_zipfile(download_path):
                        # sometimes GitHub returns HTML if blocked/rate limited
                        preview = download_path.read_bytes()[:200].decode("utf-8", errors="ignore")
                        self._log_install(f"Downloaded file not a zip. First bytes: {preview!r}")
                        raise RuntimeError("Downloaded file is not a ZIP (possible HTML/error response).")

                    unzip_dir = workdir / "unzipped"
                    unzip_dir.mkdir(parents=True, exist_ok=True)

                    with zipfile.ZipFile(download_path, "r") as z:
                        z.extractall(unzip_dir)

                    # Find Copy2.exe in extracted structure
                    candidates = list(unzip_dir.rglob("Copy2.exe"))
                    if not candidates:
                        # also accept renamed exe if user changed it
                        exes = list(unzip_dir.rglob("*.exe"))
                        raise RuntimeError(
                            f"No Copy2.exe was found inside the ZIP.\nFound EXEs: {[x.name for x in exes][:15]}"
                        )

                    # Prefer the first matching "Copy2.exe"
                    src = candidates[0]
                    staged_copy2.write_bytes(src.read_bytes())
                    self._log_install(f"Staged exe from zip: {src} -> {staged_copy2} size={staged_copy2.stat().st_size}")

                    # Optional uninstaller if present in zip
                    un_candidates = list(unzip_dir.rglob("Copy2_Uninstall.exe"))
                    if un_candidates:
                        staged_uninst.write_bytes(un_candidates[0].read_bytes())
                        self._log_install(f"Staged uninstall: {un_candidates[0]} -> {staged_uninst}")

                # Write BAT with env reset + selftest + retries + rollback
                pid = os.getpid()
                exe_name = target_exe.name
                bak_exe = target_dir / (exe_name + ".bak")
                bat = workdir / "copy2_updater.bat"

                batlog = self.data_dir / "update_bat.log"

                bat_contents = f"""@echo off
setlocal enableextensions enabledelayedexpansion

set "PID={pid}"
set "TARGET_EXE={str(target_exe)}"
set "TARGET_DIR={str(target_dir)}"
set "EXE_NAME={exe_name}"
set "BAK_EXE={str(bak_exe)}"
set "NEW_COPY2={str(staged_copy2)}"
set "NEW_UNINST={str(staged_uninst)}"
set "BATLOG={str(batlog)}"

echo {now_ts()}  BAT start>> "%BATLOG%"
echo PID=%PID%>> "%BATLOG%"
echo TARGET_EXE=%TARGET_EXE%>> "%BATLOG%"
echo TARGET_DIR=%TARGET_DIR%>> "%BATLOG%"
echo NEW_COPY2=%NEW_COPY2%>> "%BATLOG%"

:wait
for /f "tokens=2 delims=," %%A in ('tasklist /FI "PID eq %PID%" /FO CSV /NH 2^>NUL') do (
  if "%%~A"=="%PID%" (
    timeout /t 1 /nobreak >NUL
    goto wait
  )
)

pushd "%TARGET_DIR%" >NUL 2>&1
if errorlevel 1 (
  echo {now_ts()}  pushd failed>> "%BATLOG%"
  exit /b 1
)

REM Backup current exe
if exist "%BAK_EXE%" del /f /q "%BAK_EXE%" >NUL 2>&1
copy /y "%TARGET_EXE%" "%BAK_EXE%" >NUL 2>&1
if errorlevel 1 (
  echo {now_ts()}  backup failed>> "%BATLOG%"
  popd
  exit /b 1
)

REM Swap in new exe (copy, do not move)
copy /y "%NEW_COPY2%" "%TARGET_EXE%" >NUL 2>&1
if errorlevel 1 (
  echo {now_ts()}  swap failed>> "%BATLOG%"
  if exist "%BAK_EXE%" copy /y "%BAK_EXE%" "%TARGET_EXE%" >NUL 2>&1
  popd
  exit /b 1
)

REM Optional uninstaller swap if provided
if exist "%NEW_UNINST%" (
  for %%I in ("%NEW_UNINST%") do set SIZE=%%~zI
  if NOT "!SIZE!"=="0" (
    copy /y "%NEW_UNINST%" "%TARGET_DIR%\\Copy2_Uninstall.exe" >NUL 2>&1
  )
)

timeout /t 2 /nobreak >NUL

REM Critical: clear PyInstaller/Python env vars
set "PYTHONHOME="
set "PYTHONPATH="
set "_MEIPASS2="
set "PYINSTALLER_RESET_ENVIRONMENT=1"

set "TRIES=0"
:selftest
set /a TRIES+=1
echo {now_ts()}  selftest try !TRIES!>> "%BATLOG%"

"%TARGET_EXE%" --copy2-selftest >NUL 2>&1
if errorlevel 1 (
  echo {now_ts()}  selftest failed try !TRIES!>> "%BATLOG%"
  if !TRIES! LSS 5 (
    timeout /t 2 /nobreak >NUL
    goto selftest
  )

  echo {now_ts()}  selftest failed - rollback>> "%BATLOG%"
  if exist "%BAK_EXE%" copy /y "%BAK_EXE%" "%TARGET_EXE%" >NUL 2>&1
  start "" "%TARGET_EXE%"
  popd
  exit /b 1
)

echo {now_ts()}  selftest OK - launching>> "%BATLOG%"
start "" "%TARGET_EXE%"

popd
echo {now_ts()}  BAT done>> "%BATLOG%"
del "%~f0" >NUL 2>&1
endlocal
"""

                bat.write_text(bat_contents, encoding="utf-8", errors="ignore")
                self._log_install(f"Wrote updater bat: {bat}")

                # Launch bat detached
                creationflags = 0
                if hasattr(subprocess, "DETACHED_PROCESS"):
                    creationflags |= subprocess.DETACHED_PROCESS
                if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
                    creationflags |= subprocess.CREATE_NEW_PROCESS_GROUP

                subprocess.Popen(
                    ["cmd.exe", "/c", str(bat)],
                    close_fds=True,
                    creationflags=creationflags,
                    cwd=str(workdir),
                )

                self.status_var.set(f"Applying update and restarting... — {now_ts()}")
                self.after(150, self._on_close_for_update)
                return

            except Exception as e:
                self._log_install(f"ERROR: {repr(e)}")

                def ui_err():
                    messagebox.showerror(
                        APP_NAME,
                        f"Update failed:\n{e}\n\nLog:\n{self.data_dir / 'update_install.log'}",
                    )
                    self.status_var.set(f"Update failed — {now_ts()}")

                self.after(0, ui_err)

        threading.Thread(target=worker, daemon=True).start()

    def _on_close_for_update(self):
        # Close without unsaved edits blocking (updater already staged)
        try:
            if self._poll_job is not None:
                self.after_cancel(self._poll_job)
        except Exception:
            pass
        try:
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


# -----------------------------
# ttkbootstrap variant
# -----------------------------
if USE_TTKB:

    class Copy2App(tb.Window, Copy2AppBase):
        def __init__(self, *args, **kwargs):
            tmp_settings = Settings.from_dict(safe_json_load(Path(user_data_dir(APP_ID, VENDOR)) / "config.json", {}))
            theme = tmp_settings.theme if tmp_settings.theme else "flatly"
            super().__init__(themename=theme, *args, **kwargs)
            self._apply_app_icon_to_window()
            self._active_theme = theme

            self.title(APP_NAME)
            self._init_state()
            self._apply_window_defaults()

            # Build UI first. Locking is UI-only and applied as an overlay after launch.
            self._build_ui()
            self._bind_shortcuts()
            self._bind_inactivity_listeners()
            self._build_context_menu()
            self._refresh_list(select_last=True)

            # Optional: global hotkeys & sync
            self._register_global_hotkeys()
            self._start_sync_job()

            # Start clipboard engine unconditionally (must continue while locked)
            self.after(250, self._poll_clipboard)
            self.protocol("WM_DELETE_WINDOW", self._on_close)

            # Apply startup lock overlay (UI-only; engine keeps running)
            self.after(0, self._apply_startup_lock_overlay)

            # Start inactivity timer (only does something if enabled)
            self.after(0, self._schedule_inactivity_lock)

            # Auto-check updates on launch (non-blocking)
            if self.settings.check_updates_on_launch:
                self.after(900, lambda: self._check_updates_async(prompt_if_new=True))


        def _apply_window_defaults(self):
            self.minsize(1100, 720)
            try:
                self.tk.call("tk", "scaling", 1.2)
            except Exception:
                pass

        def _build_ui(self):
            root = tb.Frame(self, padding=12)
            root.pack(fill=BOTH, expand=True)

            # Top bar
            top = tb.Frame(root)
            top.pack(fill=X, pady=(0, 10))

            title_box = tb.Frame(top)
            title_box.pack(side=LEFT, padx=(0, 14))
            tb.Label(title_box, text="Copy 2.0", font=("Segoe UI", 18, "bold")).pack(anchor="w")
            self.capture_state_var = tk.StringVar(value="Capturing")
            tb.Label(title_box, textvariable=self.capture_state_var, font=("Segoe UI", 10)).pack(anchor="w")

            # Search cluster
            search_box = tb.Frame(top)
            search_box.pack(side=LEFT, fill=X, expand=True)

            tb.Label(search_box, text="Search").pack(side=LEFT, padx=(0, 8))
            self.search_var = tk.StringVar(value="")
            self.search_entry = tb.Entry(search_box, textvariable=self.search_var)
            self.search_entry.pack(side=LEFT, fill=X, expand=True)
            try:
                self.search_entry.bind('<KeyRelease>', self._search_live)
            except Exception:
                pass

            btn_find = tb.Button(search_box, text="Find", command=self._search, bootstyle=PRIMARY)
            btn_find.pack(side=LEFT, padx=(10, 0))
            ToolTip(btn_find, "Find (Enter)")

            btn_prev = tb.Button(search_box, text="Prev", command=lambda: self._jump_match(-1), bootstyle=SECONDARY)
            btn_prev.pack(side=LEFT, padx=(8, 0))

            btn_next = tb.Button(search_box, text="Next", command=lambda: self._jump_match(1), bootstyle=SECONDARY)
            btn_next.pack(side=LEFT, padx=(8, 0))

            # Spacer to ensure Pause isn't too close
            tb.Frame(top, width=18).pack(side=LEFT)

            # Right actions
            action_box = tb.Frame(top)
            action_box.pack(side=RIGHT)

            self.pause_var = tk.BooleanVar(value=self.paused)
            pause_btn = tb.Checkbutton(
                action_box,
                text="Pause",
                variable=self.pause_var,
                command=self._toggle_pause,
                bootstyle="round-toggle",
            )
            pause_btn.pack(side=LEFT, padx=(10, 18))
            ToolTip(pause_btn, "Pause capturing")

            btn_updates = tb.Button(action_box, text="Check Updates", command=lambda: self._check_updates_async(True), bootstyle=INFO)
            btn_updates.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_updates, "Check for updates now")

            btn_export = tb.Button(action_box, text="Export", command=self._export, bootstyle=OUTLINE)
            btn_export.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_export, "Export (Ctrl+E)")

            btn_import = tb.Button(action_box, text="Import", command=self._import, bootstyle=OUTLINE)
            btn_import.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_import, "Import (Ctrl+I)")

            btn_quick = tb.Button(action_box, text="Quick Paste", command=self._open_quick_paste, bootstyle=PRIMARY)
            btn_quick.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_quick, "Quick Paste palette")

            btn_sync_now = tb.Button(action_box, text="Sync Now", command=self._sync_now, bootstyle=OUTLINE)
            btn_sync_now.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_sync_now, "Manual sync (if enabled)")

            btn_settings = tb.Button(action_box, text="Settings", command=self._open_settings, bootstyle=SECONDARY)
            btn_settings.pack(side=LEFT)
            ToolTip(btn_settings, "Settings (includes hotkeys list)")

            tb.Separator(root).pack(fill=X, pady=(0, 10))

            # Body split (user-resizable)
            self.paned = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, showhandle=True, opaqueresize=True)
            self.paned.pack(fill=BOTH, expand=True)

            # Left pane (History)
            left = tb.Labelframe(self.paned, text="History", padding=10)
            # Right pane (Preview)
            right = tb.Labelframe(self.paned, text="Preview (Editable)", padding=10)

            self.paned.add(left, minsize=340)
            self.paned.add(right, minsize=520)

            def _apply_sash():
                try:
                    self.paned.sash_place(0, int(self.settings.pane_sash), 0)
                except Exception:
                    pass

            def _store_sash(_evt=None):
                # Persist the left pane width (px), clamped to current window constraints.
                try:
                    x, _y = self.paned.sash_coord(0)
                    total = int(self.paned.winfo_width() or 0)
                    min_left = 340
                    min_right = 520
                    if total > 0:
                        max_left = max(min_left, total - min_right)
                        x = max(min_left, min(int(x), max_left))
                    self.settings.pane_sash = int(x)
                except Exception:
                    pass

            self._sash_dragging = False
            self._sash_apply_guard = False

            def _store_sash_motion(_evt=None):
                self._sash_dragging = True
                _store_sash(_evt)

            def _store_sash_release(_evt=None):
                _store_sash(_evt)
                self._sash_dragging = False

            def _keep_sash_on_resize(_evt=None):
                # Keep the left pane width fixed (px) when the main window is resized.
                if self._sash_dragging or self._sash_apply_guard:
                    return
                try:
                    total = int(self.paned.winfo_width() or 0)
                    if total <= 0:
                        return
                    min_left = 340
                    min_right = 520
                    desired = int(self.settings.pane_sash)
                    max_left = max(min_left, total - min_right)
                    desired = max(min_left, min(desired, max_left))
                    cur, _ = self.paned.sash_coord(0)
                    if abs(int(cur) - desired) <= 2:
                        return
                    self._sash_apply_guard = True

                    def _do():
                        try:
                            self.paned.sash_place(0, desired, 0)
                        except Exception:
                            pass
                        self._sash_apply_guard = False

                    self.after_idle(_do)
                except Exception:
                    pass

            self.after(80, _apply_sash)
            self.paned.bind("<Configure>", _keep_sash_on_resize)
            self.paned.bind("<ButtonRelease-1>", _store_sash_release)
            self.paned.bind("<B1-Motion>", _store_sash_motion)

            left.rowconfigure(2, weight=1)
            left.columnconfigure(0, weight=1)

            # Filters row (responsive): when the History pane is narrowed, action buttons stack underneath
            filters = tb.Frame(left)
            filters.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            filters.columnconfigure(0, weight=1)
            filters.columnconfigure(1, weight=0)

            filters_main = tb.Frame(filters)
            filters_main.grid(row=0, column=0, sticky="ew")

            filters_actions = tb.Frame(filters)
            filters_actions.grid(row=0, column=1, sticky="e")

            self.filter_var = tk.StringVar(value="all")
            tb.Radiobutton(filters_main, text="All", value="all", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 8))
            tb.Radiobutton(filters_main, text="Favorites", value="fav", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 8))
            tb.Radiobutton(filters_main, text="Pinned", value="pin", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 12))
            tb.Radiobutton(filters_main, text="Images", value="img", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 12))

            tb.Label(filters_main, text="Tag").pack(side=LEFT)
            self.tag_filter_var = tk.StringVar(value="")
            self.tag_combo = tb.Combobox(filters_main, textvariable=self.tag_filter_var, values=[""], width=18, state="readonly")
            self.tag_combo.pack(side=LEFT, padx=(8, 0))
            self.tag_combo.bind("<<ComboboxSelected>>", lambda e: self._on_tag_filter_change())

            btn_clean = tb.Button(filters_actions, text="Clean", command=self._clean_keep_favorites, bootstyle=DANGER)
            btn_clean.pack()
            ToolTip(btn_clean, "Clean (keeps favorites/pins)  (Ctrl+L)")

            self._filters_compact = False

            def _relayout_filters(_evt=None):
                # If the History pane becomes too narrow, stack the action button(s) underneath.
                try:
                    w = int(filters.winfo_width() or 0)
                except Exception:
                    w = 0

                threshold = 620  # tweak point where the Clean button begins to clip on small panes
                if w and w < threshold and not self._filters_compact:
                    self._filters_compact = True
                    filters_actions.grid_configure(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
                    try:
                        btn_clean.pack_forget()
                    except Exception:
                        pass
                    btn_clean.pack(fill=X)
                elif w and w >= threshold and self._filters_compact:
                    self._filters_compact = False
                    filters_actions.grid_configure(row=0, column=1, columnspan=1, sticky="e", pady=(0, 0))
                    try:
                        btn_clean.pack_forget()
                    except Exception:
                        pass
                    btn_clean.pack()

            filters.bind("<Configure>", _relayout_filters)
            self.after(120, _relayout_filters)

            self.listbox = tk.Listbox(left, activestyle="dotbox", exportselection=False, selectmode=tk.EXTENDED)
            # Improve readability (Listbox is not ttk-themed by default)
            try:
                ui_font = ("Segoe UI", 11)
                colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
                if colors is not None:
                    self.listbox.configure(
                        font=ui_font,
                        bg=colors.bg,
                        fg=colors.fg,
                        selectbackground=colors.primary,
                        selectforeground=colors.light if hasattr(colors, 'light') else colors.fg,
                    )
                else:
                    self.listbox.configure(font=ui_font)
            except Exception:
                pass

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

            btn_copy = tb.Button(left_actions, text="Copy", command=self._copy_selected, bootstyle=SUCCESS)
            btn_copy.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            ToolTip(btn_copy, "Copy Preview (Ctrl+C)")

            btn_del = tb.Button(left_actions, text="Delete", command=self._delete_selected, bootstyle=WARNING)
            btn_del.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            ToolTip(btn_del, "Delete (Del)")

            btn_fav = tb.Button(left_actions, text="Fav / Unfav", command=self._toggle_favorite_selected, bootstyle=INFO)
            btn_fav.grid(row=0, column=2, sticky="ew", padx=(0, 8))

            btn_combine = tb.Button(left_actions, text="Combine", command=self._combine_selected, bootstyle=PRIMARY)
            btn_combine.grid(row=0, column=3, sticky="ew")

            btn_pin = tb.Button(left_actions, text="Pin / Unpin", command=self._toggle_pin_selected, bootstyle=SECONDARY)
            btn_pin.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(8, 0))

            btn_tags = tb.Button(left_actions, text="Tags", command=self._open_tags_dialog, bootstyle=SECONDARY)
            btn_tags.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(8, 0))

            btn_exp = tb.Button(left_actions, text="Expiry", command=self._set_expiry_selected, bootstyle=SECONDARY)
            btn_exp.grid(row=1, column=2, sticky="ew", padx=(0, 8), pady=(8, 0))

            btn_qp = tb.Button(left_actions, text="Quick Paste", command=self._open_quick_paste, bootstyle=PRIMARY)
            btn_qp.grid(row=1, column=3, sticky="ew", pady=(8, 0))

            # Right pane (Preview)
            # (already created as a paned window pane)
            right.rowconfigure(1, weight=1)
            right.columnconfigure(0, weight=1)

            preview_actions = tb.Frame(right)
            preview_actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            preview_actions.columnconfigure(0, weight=1)

            self.reverse_var = tk.BooleanVar(value=False)
            rev_btn = tb.Checkbutton(
                preview_actions,
                text="Reverse-lines copy",
                variable=self.reverse_var,
                command=self._toggle_reverse_lines,
                bootstyle="round-toggle",
            )
            rev_btn.pack(side=LEFT)
            ToolTip(rev_btn, "Reverse lines in Preview and Copy")

            self.preview_dirty_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.preview_dirty_var).pack(side=LEFT, padx=(10, 0))

            self.revert_btn = tb.Button(preview_actions, text="Revert", command=self._revert_preview_edits, bootstyle=WARNING)
            self.revert_btn.pack(side=RIGHT, padx=(0, 8))
            ToolTip(self.revert_btn, "Revert (Esc)")

            self.save_btn = tb.Button(preview_actions, text="Save Edit", command=self._save_preview_edits, bootstyle=SUCCESS)
            self.save_btn.pack(side=RIGHT, padx=(0, 8))
            ToolTip(self.save_btn, "Save (Ctrl+S)")

            # Format tools dropdown (Option A: fast menu)
            btn_format = tb.Button(preview_actions, text='Format Tools', bootstyle=INFO)
            btn_format.pack(side=RIGHT, padx=(0, 18))
            ToolTip(btn_format, 'Formatting tools for the Preview text')

            # Build menu once
            self._format_menu = tk.Menu(self, tearoff=0)
            self._format_menu.add_command(label='Strip hidden characters', command=self._fmt_strip_hidden)
            self._format_menu.add_command(label='Remove blank lines', command=self._fmt_remove_blank_lines)
            self._format_menu.add_command(label='Collapse multiple spaces', command=self._fmt_collapse_spaces)
            self._format_menu.add_command(label='Trim each line (leading/trailing)', command=self._fmt_trim_each_line)
            self._format_menu.add_command(label='Strip trailing whitespace', command=self._fmt_strip_trailing_ws)
            self._format_menu.add_command(label='Normalize line endings (CRLF → LF)', command=self._fmt_normalize_line_endings)
            self._format_menu.add_separator()
            self._format_menu.add_command(label='Strip hidden + remove blank lines', command=self._fmt_strip_hidden_and_blanks)

            self._format_menu.add_separator()
            self._format_menu.add_command(label='Paste plain text into Preview', command=self._fmt_paste_plain_text)
            self._format_menu.add_command(label='Copy Preview as plain text', command=self._fmt_copy_preview_plain)
            self._format_menu.add_command(label='Open URL(s) found in Preview', command=self._fmt_open_urls_in_preview)

            btn_format.configure(command=lambda w=btn_format: self._show_format_menu(w))

            btn_copy_prev = tb.Button(preview_actions, text='Copy Preview', command=self._copy_selected, bootstyle=SUCCESS)
            btn_copy_prev.pack(side=RIGHT, padx=(0, 18))
            ToolTip(btn_copy_prev, 'Copy (Ctrl+C)')

            # Preview container supports both text preview (with line numbers) and image preview.
            self.preview_container = tb.Frame(right)
            self.preview_container.grid(row=1, column=0, columnspan=2, sticky="nsew")
            self.preview_container.rowconfigure(0, weight=1)
            self.preview_container.columnconfigure(0, weight=1)

            # --- Text preview (with line numbers gutter)
            self.preview_text_frame = tb.Frame(self.preview_container)
            self.preview_text_frame.grid(row=0, column=0, sticky="nsew")
            self.preview_text_frame.rowconfigure(0, weight=1)
            self.preview_text_frame.columnconfigure(1, weight=1)

            self.preview_gutter = tk.Text(self.preview_text_frame, width=6, padx=6, pady=6, wrap='none', state='disabled')
            self.preview_gutter.grid(row=0, column=0, sticky='ns')
            try:
                self.preview_gutter.configure(takefocus=0, cursor='arrow')
            except Exception:
                pass

            self.preview = tk.Text(self.preview_text_frame, wrap="word", undo=True)
            # Improve readability (Text is not ttk-themed by default)
            try:
                prev_font = ("Segoe UI", 11)
                colors = getattr(self, 'style', None).colors if getattr(self, 'style', None) is not None else None
                if colors is not None:
                    self.preview.configure(
                        font=prev_font,
                        bg=colors.bg,
                        fg=colors.fg,
                        insertbackground=colors.fg,
                    )
                    # Make gutter visually subtle
                    try:
                        self.preview_gutter.configure(font=prev_font, bg=colors.bg, fg=colors.fg)
                    except Exception:
                        pass
                else:
                    self.preview.configure(font=prev_font)
                    try:
                        self.preview_gutter.configure(font=prev_font)
                    except Exception:
                        pass
            except Exception:
                pass

            self.preview.grid(row=0, column=1, sticky="nsew")

            def _preview_yview(*args):
                try:
                    self.preview.yview(*args)
                    self.preview_gutter.yview(*args)
                except Exception:
                    pass

            sb2 = tb.Scrollbar(self.preview_text_frame, orient="vertical", command=_preview_yview)
            sb2.grid(row=0, column=2, sticky="ns")

            def _yscroll(first, last):
                try:
                    sb2.set(first, last)
                except Exception:
                    pass
                try:
                    self.preview_gutter.yview_moveto(first)
                except Exception:
                    pass

            self.preview.configure(yscrollcommand=_yscroll)
            # Note: do not set yscrollcommand on the gutter to avoid recursion;
            # we drive its scroll position from the main Preview Text widget.

            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)
            self.preview.bind("<MouseWheel>", lambda e: (self.preview.yview_scroll(int(-1*(e.delta/120)), "units"), self.preview_gutter.yview_scroll(int(-1*(e.delta/120)), "units"), 'break'))

            # --- Image preview
            self.preview_image_frame = tb.Frame(self.preview_container)
            self.preview_image_frame.grid(row=0, column=0, sticky='nsew')
            self.preview_image_frame.rowconfigure(1, weight=1)
            self.preview_image_frame.columnconfigure(0, weight=1)

            self.image_preview_title = tb.Label(self.preview_image_frame, text='', anchor='w')
            self.image_preview_title.grid(row=0, column=0, sticky='ew', padx=6, pady=(6, 0))
            self.image_preview_label = tb.Label(self.preview_image_frame)
            self.image_preview_label.grid(row=1, column=0, sticky='nsew', padx=6, pady=6)

            # Start with text preview visible
            try:
                self.preview_image_frame.grid_remove()
            except Exception:
                pass

            # Bottom status bar
            bottom = tb.Frame(root)
            bottom.pack(fill=X, pady=(10, 0))
            self.status_var = tk.StringVar(value='')
            tb.Label(bottom, textvariable=self.status_var).pack(side=LEFT)

            self._update_preview_dirty_ui()

            # Apply theme colors to tk widgets and initialize status bar
            try:
                self._apply_theme_to_tk_widgets()
            except Exception:
                pass
            try:
                self._update_status_bar()
            except Exception:
                pass

        def _bind_shortcuts(self):
            self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
            self.bind("<Return>", lambda e: (self._search(), "break"))
            self.bind("<Control-c>", lambda e: (self._copy_selected(), "break"))
            self.bind("<Delete>", lambda e: (self._delete_selected(), "break"))
            self.bind("<Control-e>", lambda e: (self._export(), "break"))
            self.bind("<Control-i>", lambda e: (self._import(), "break"))
            self.bind("<Control-l>", lambda e: (self._clean_keep_favorites(), "break"))
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))

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

            self._unregister_global_hotkeys()
            self._stop_sync_job()
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            self.destroy()


else:
    # ttk fallback (minimal; modern UI requires ttkbootstrap)
    class Copy2App(tk.Tk, Copy2AppBase):
        def __init__(self):
            super().__init__()
            # Apply window/taskbar icon (best-effort; requires Image 4.png next to the exe or bundled)
            self._apply_app_icon_to_window()
            self.title(APP_NAME)
            self._init_state()
            self.minsize(1100, 720)

            # Build UI first. Locking is UI-only and applied as an overlay after launch.
            self._build_ui()
            self._bind_shortcuts()
            self._bind_inactivity_listeners()
            self._build_context_menu()
            self._refresh_list(select_last=True)

            # Optional: global hotkeys (Windows; requires 'keyboard')
            self._register_global_hotkeys()

            # Optional: sync folder (basic pull/push)
            self._start_sync_job()

            # Start clipboard engine unconditionally (must continue while locked)
            self.after(250, self._poll_clipboard)
            self.protocol("WM_DELETE_WINDOW", self._on_close)

            # Apply startup lock overlay (UI-only; engine keeps running)
            self.after(0, self._apply_startup_lock_overlay)

            # Start inactivity timer (only does something if enabled)
            self.after(0, self._schedule_inactivity_lock)

            if self.settings.check_updates_on_launch:
                self.after(900, lambda: self._check_updates_async(prompt_if_new=True))


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
            self.reverse_var = tk.BooleanVar(value=False)
            self.status_var = tk.StringVar(value="")

            self.preview_dirty_var = tk.StringVar(value="")
            self.save_btn = None
            self.revert_btn = None

        def _bind_shortcuts(self):
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))
            self.bind("<Control-l>", lambda e: (self._clean_keep_favorites(), "break"))

        def _on_close(self):
            try:
                if self._poll_job is not None:
                    self.after_cancel(self._poll_job)
            except Exception:
                pass
            self._unregister_global_hotkeys()
            self._stop_sync_job()
            self._persist()
            try:
                self._schedule_inactivity_lock()
            except Exception:
                pass
            self.destroy()


def main():
    # Selftest for updater BAT: if this runs, Python DLLs loaded successfully.
    if "--copy2-selftest" in sys.argv:
        sys.exit(0)

    # Prevent duplicate instances (fixes double-startup and PIN lock crashes)
    try:
        if os.name == 'nt':
            import ctypes
            from ctypes import wintypes
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            CreateMutexW = kernel32.CreateMutexW
            CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
            CreateMutexW.restype = wintypes.HANDLE
            GetLastError = kernel32.GetLastError
            ERROR_ALREADY_EXISTS = 183
            h = CreateMutexW(None, False, 'Global\\Copy2SingleInstanceMutex')
            if h and GetLastError() == ERROR_ALREADY_EXISTS:
                # Another instance is already running
                return
    except Exception:
        pass

    app = Copy2App() if not USE_TTKB else Copy2App()
    app.mainloop()


if __name__ == "__main__":
    main()
