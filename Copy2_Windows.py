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
import shutil
import webbrowser
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
APP_VERSION = "1.0.6"  # <-- keep this in sync with your build/tag

# History indicator icons (you can change these to any emoji you like)
PIN_ICON = "📌"
FAV_ICON = "⭐"
TAG_ICON = "🏷️"

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

    # File-based sync (point to a cloud-synced folder like OneDrive/Dropbox)
    sync_enabled: bool = False
    sync_folder: str = ""
    sync_interval_sec: int = 10

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

        s.sync_enabled = bool(d.get("sync_enabled", False))
        s.sync_folder = str(d.get("sync_folder", "")).strip()
        s.sync_interval_sec = int(d.get("sync_interval_sec", 10))

        s.max_history = max(5, min(HARD_MAX_HISTORY, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        s.pane_sash = max(220, min(900, s.pane_sash))
        s.sync_interval_sec = max(3, min(300, s.sync_interval_sec))
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

    def _init_state(self):
        self.data_dir = Path(user_data_dir(APP_ID, VENDOR))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"
        self.pins_path = self.data_dir / "pins.json"
        self.tags_path = self.data_dir / "tags.json"
        self.tag_colors_path = self.data_dir / "tag_colors.json"
        self.expiry_path = self.data_dir / "expiry.json"

        # Sync folder state
        self._sync_mtimes = {}
        self._sync_job = None


        self.log_update_check = self.data_dir / "update_check.log"
        self.log_update_install = self.data_dir / "update_install.log"

        self.settings = Settings.from_dict(safe_json_load(self.settings_path, {}))

        # Normalize stored favorites to unique list
        favs = safe_json_load(self.favs_path, [])
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
        pins = safe_json_load(self.pins_path, [])
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
        tags = safe_json_load(self.tags_path, {})
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
        tc = safe_json_load(self.tag_colors_path, {})
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
        exp = safe_json_load(self.expiry_path, {})
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

        hist = safe_json_load(self.history_path, [])
        if isinstance(hist, list):
            hist = [x for x in hist if isinstance(x, str)]
        else:
            hist = []

        self.history = deque(hist, maxlen=self.settings.max_history)

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

        # Ensure favorites remain present in history (optional)
        self._ensure_favorites_present()

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

    def _persist(self):
        # Always persist settings (UI/theme/hotkeys should survive restarts).
        safe_json_save(self.settings_path, asdict(self.settings))
        if self.settings.session_only:
            return
        safe_json_save(self.history_path, list(self.history))
        safe_json_save(self.favs_path, list(self.favorites))
        safe_json_save(self.pins_path, list(self.pins))
        safe_json_save(self.tags_path, self.tags)
        safe_json_save(self.tag_colors_path, getattr(self, "tag_colors", {}))
        safe_json_save(self.expiry_path, self.expiry)


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
        if not self._clipboard_set_text(item):
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

        keep = set(self.favorites) | set(self.pins)

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
        try:
            # Expiry housekeeping (throttled)
            try:
                if time.time() - getattr(self, "_last_expiry_purge", 0.0) > 30:
                    self._purge_expired(silent=True)
                    self._last_expiry_purge = time.time()
            except Exception:
                pass
            if not self.paused:
                text = self._clipboard_get_text()
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
            # Favorites exceed capacity -> block adding the new item
            # Keep history unchanged, do not overwrite favorites.
            self._notify_favorites_blocking()

            # If they are already at hard cap, give stronger warning.
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
            return len(expired)
        except Exception:
            return 0

    # -----------------------------
    # List/view helpers
    # -----------------------------
    def _format_list_item(self, item: str) -> str:
        one = re.sub(r"\s+", " ", item).strip()
        if len(one) > 90:
            one = one[:87] + "..."

        # Use emoji indicators (customizable via PIN_ICON/FAV_ICON/TAG_ICON)
        pin_mark = PIN_ICON if item in self.pins else " "
        fav_mark = FAV_ICON if item in self.favorites else " "
        tag_mark = TAG_ICON if self.tags.get(item) else " "

        prefix = f"{pin_mark}{fav_mark}{tag_mark} "
        return prefix + one

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
        # Preserve current selection (by value) across refreshes unless select_last is requested
        preserve_texts = []
        try:
            for i in self.listbox.curselection():
                if 0 <= i < len(getattr(self, 'view_items', [])):
                    preserve_texts.append(self.view_items[i])
        except Exception:
            preserve_texts = []

        if not preserve_texts:
            try:
                preserve_texts = list(getattr(self, '_last_selected_texts', []) or [])
            except Exception:
                preserve_texts = []
            if not preserve_texts and getattr(self, '_selected_item_text', None):
                preserve_texts = [self._selected_item_text]

        items = list(self.history)
        f = self._current_filter()
        if f == "fav":
            items = [x for x in items if x in self.favorites]
        elif f == "pin":
            items = [x for x in items if x in self.pins]
        elif f == "tag":
            tag = ""
            try:
                tag = str(self.tag_filter_var.get()).strip()
            except Exception:
                tag = ""
            if tag:
                items = [x for x in items if tag in self.tags.get(x, [])]

        # Pinned items always float to the top (within the current filter)
        if items:
            pins = [x for x in items if x in self.pins]
            rest = [x for x in items if x not in self.pins]
            items = pins + rest

        self.view_items = items

        self.listbox.delete(0, tk.END)
        for idx, item in enumerate(self.view_items):
            self.listbox.insert(tk.END, self._format_list_item(item))

            # Tag color (first matching tag with a configured color)
            try:
                for tg in self.tags.get(item, []):
                    col = getattr(self, 'tag_colors', {}).get(tg)
                    if isinstance(col, str) and col.strip():
                        self.listbox.itemconfig(idx, fg=col.strip())
                        break
            except Exception:
                pass

        # Reset selection tracking
        self._prev_sel_set = set()
        self._sel_order = []

        if select_last and self.view_items:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(tk.END)
            self.listbox.see(tk.END)
            self._on_select()
        else:
            # Restore previous selection (do not trigger _on_select unless selection is empty)
            if preserve_texts:
                try:
                    self.listbox.selection_clear(0, tk.END)
                    restored = []
                    for t in preserve_texts:
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

        try:
            self._update_status_bar()
        except Exception:
            self.status_var.set(
                f"v{APP_VERSION}   Items: {len(self.history)}   Favorites: {len(self.favorites)}   Pins: {len(self.pins)}   Data: {self.data_dir}"
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

        # Keep reverse-lines checkbox in sync with the selected item
        try:
            self.reverse_var.set(getattr(self, '_reverse_item_text', None) == t)
        except Exception:
            pass
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
        out = self.preview.get("1.0", tk.END).rstrip("\n")
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
        """
        Clean action: remove all NON-favorites, keep favorites.
        This matches your requirement: favorites stay behind after cleaning.
        """
        if not messagebox.askyesno(APP_NAME, "Clean history and keep Favorites/Pins only?"):
            return

        keep = set(self.favorites) | set(self.pins)
        kept = [x for x in self.history if x in keep]

        self.history = deque(kept, maxlen=self.settings.max_history)

        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._refresh_list(select_last=True)
        self._persist()
        self.status_var.set(f"Cleaned (kept favorites/pins) — {now_ts()}")

    def _toggle_favorite_selected(self):
        t = self._get_selected_text()
        if not t:
            return

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]
        else:
            self.favorites.append(t)

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
            self._clipboard_set_text(t)
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
                self._clipboard_set_text(t)
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
        """Live search feedback while typing (does not jump selection)."""
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

        matches = []
        ql = q.lower()
        for item in list(self.history):
            try:
                if ql in item.lower() or self._fuzzy_match(q, item):
                    matches.append(item)
            except Exception:
                pass
        self.search_matches = matches
        self.search_index = 0

        # Highlight in preview (all occurrences of the raw query)
        try:
            self._highlight_query_in_preview(q)
        except Exception:
            pass

        try:
            self._set_status_note(f"Search: {len(matches)} match(es)")
        except Exception:
            pass

    def _search(self):
        q = self.search_var.get().strip()
        self.search_query = q
        self.search_matches = []
        self.search_index = 0

        if not q:
            self.status_var.set("Search cleared.")
            try:
                self.preview.tag_remove("match", "1.0", tk.END)
            except Exception:
                pass
            return

        # Match against full history using substring OR loose fuzzy (in-order chars)
        ql = q.lower()
        for item in list(self.history):
            if ql in item.lower() or self._fuzzy_match(q, item):
                self.search_matches.append(item)

        if not self.search_matches:
            self.status_var.set(f"No matches for: {q}")
            try:
                self.preview.tag_remove("match", "1.0", tk.END)
            except Exception:
                pass
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
    # Import / Export (verified, stable)
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
            "exported_at": now_ts(),
            "app_version": APP_VERSION,
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
            history = data.get("history", [])
            favorites = data.get("favorites", [])

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

            # merge favorites (unique)
            if isinstance(favorites, list):
                for x in favorites:
                    if isinstance(x, str) and x not in self.favorites:
                        self.favorites.append(x)

            cap = self.settings.max_history
            pruned, ok = self._prune_preserving_favorites(out, cap)
            if not ok:
                # If favorites exceed cap, keep only favorites that are present in the pruned list
                self._notify_favorites_blocking()

            self.history = deque(pruned[-cap:], maxlen=cap)

            self._refresh_list(select_last=True)
            self._persist()
            self.status_var.set(f"Imported — {path}")

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")

    # -----------------------------
    # Settings dialog (+ Controls section)
    # -----------------------------

    # -----------------------------
    # Settings dialog (General / Hotkeys / Sync / Help)
    # -----------------------------
    def _open_settings(self, initial_tab: str = "General"):
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
        tab_help = ttk.Frame(nb)

        nb.add(tab_general, text='General')
        nb.add(tab_hotkeys, text='Hotkeys')
        nb.add(tab_sync, text='Sync')
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

        ttk.Label(g, text="Theme:").grid(row=4, column=0, sticky='w', padx=10, pady=(10,6))
        cb_theme = ttk.Combobox(g, textvariable=theme_var, values=themes, width=22, state='readonly')
        cb_theme.grid(row=4, column=1, sticky='w', padx=10, pady=(10,6))
        if not USE_TTKB:
            ttk.Label(g, text="(Install ttkbootstrap to enable theme switching)").grid(row=5, column=0, columnspan=2, sticky='w', padx=10, pady=(6,0))

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

        def save():
            try:
                mh = int(max_var.get().strip())
                pm = int(poll_var.get().strip())
            except Exception:
                messagebox.showerror(APP_NAME, 'Please enter valid integers for Max history and Poll interval.')
                return

            mh = max(5, min(HARD_MAX_HISTORY, mh))
            pm = max(100, min(5000, pm))

            if mh >= HARD_MAX_HISTORY:
                self._notify_hard_cap_reached()

            # Save settings
            self.settings.max_history = mh
            self.settings.poll_ms = pm
            self.settings.session_only = bool(sess_var.get())
            self.settings.check_updates_on_launch = bool(upd_var.get())

            if USE_TTKB:
                new_theme = str(theme_var.get()).strip() or 'flatly'
                self.settings.theme = new_theme
                _apply_theme_now(new_theme)

            self.settings.enable_global_hotkeys = bool(hk_en_var.get())
            self.settings.hotkey_quick_paste = str(hk_qp_var.get()).strip() or self.settings.hotkey_quick_paste
            self.settings.hotkey_paste_last = str(hk_pl_var.get()).strip() or self.settings.hotkey_paste_last

            self.settings.sync_enabled = bool(sync_en_var.get())
            self.settings.sync_folder = str(sync_folder_var.get()).strip()
            try:
                self.settings.sync_interval_sec = max(3, min(300, int(str(sync_int_var.get()).strip())))
            except Exception:
                self.settings.sync_interval_sec = max(3, min(300, int(getattr(self.settings, 'sync_interval_sec', 10) or 10)))

            # Apply history capacity (preserve favorites/pins)
            items = list(self.history)
            pruned, ok = self._prune_preserving_favorites(items, mh)
            if not ok:
                self._notify_favorites_blocking()
            self.history = deque(pruned[-mh:], maxlen=mh)

            self._refresh_list(select_last=True)
            self._persist()

            # Apply runtime features
            self._register_global_hotkeys()
            self._start_sync_job()

            dlg.destroy()

        ttk.Button(bottom, text='Cancel', command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(bottom, text='Save', command=save).pack(side=tk.RIGHT, padx=(0,10))

        # Select initial tab
        tab_map = {'general': 0, 'hotkeys': 1, 'sync': 2, 'help': 3}
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

            # Optional: global hotkeys & sync
            self._register_global_hotkeys()
            self._start_sync_job()

            self.after(250, self._poll_clipboard)
            self.protocol("WM_DELETE_WINDOW", self._on_close)

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

            filters = tb.Frame(left)
            filters.grid(row=0, column=0, sticky="ew", pady=(0, 10))

            self.filter_var = tk.StringVar(value="all")
            tb.Radiobutton(filters, text="All", value="all", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 8))
            tb.Radiobutton(filters, text="Favorites", value="fav", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 8))
            tb.Radiobutton(filters, text="Pinned", value="pin", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 12))

            tb.Label(filters, text="Tag").pack(side=LEFT)
            self.tag_filter_var = tk.StringVar(value="")
            self.tag_combo = tb.Combobox(filters, textvariable=self.tag_filter_var, values=[""], width=18, state="readonly")
            self.tag_combo.pack(side=LEFT, padx=(8, 0))
            self.tag_combo.bind("<<ComboboxSelected>>", lambda e: self._on_tag_filter_change())

            btn_clean = tb.Button(filters, text="Clean", command=self._clean_keep_favorites, bootstyle=DANGER)
            btn_clean.pack(side=RIGHT)
            ToolTip(btn_clean, "Clean (keeps favorites/pins)  (Ctrl+L)")

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

            self.preview = tk.Text(right, wrap="word", undo=True)
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
                else:
                    self.preview.configure(font=prev_font)
            except Exception:
                pass

            self.preview.grid(row=1, column=0, sticky="nsew")
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            sb2 = tb.Scrollbar(right, orient="vertical", command=self.preview.yview)
            sb2.grid(row=1, column=1, sticky="ns")
            self.preview.configure(yscrollcommand=sb2.set)

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
            self.destroy()


else:
    # ttk fallback (minimal; modern UI requires ttkbootstrap)
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

            # Optional: global hotkeys (Windows; requires 'keyboard')
            self._register_global_hotkeys()
            
            # Optional: sync folder (basic pull/push)
            self._start_sync_job()

            self.after(250, self._poll_clipboard)
            self.protocol("WM_DELETE_WINDOW", self._on_close)

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
            self.destroy()


def main():
    # Selftest for updater BAT: if this runs, Python DLLs loaded successfully.
    if "--copy2-selftest" in sys.argv:
        sys.exit(0)

    app = Copy2App() if not USE_TTKB else Copy2App()
    app.mainloop()


if __name__ == "__main__":
    main()
