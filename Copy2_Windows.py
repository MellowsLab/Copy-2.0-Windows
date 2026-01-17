"""
Copy 2.0 (Windows) — Modern UI Edition (Improved)

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
import difflib
import ctypes
import sys
import time
import zipfile
import tempfile
import threading
import subprocess
import webbrowser
from collections import deque
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

import pyperclip
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

# -----------------------------
# App constants
# -----------------------------
APP_NAME = "Copy 2.0"
APP_ID = "copy2"
VENDOR = "MellowsLab"

# App version (display + update compare)
APP_VERSION = "1.0.5"  # <-- keep this in sync with your build/tag

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

    @staticmethod
    def from_dict(d: dict) -> "Settings":
        s = Settings()
        s.max_history = int(d.get("max_history", DEFAULT_MAX_HISTORY))
        s.poll_ms = int(d.get("poll_ms", DEFAULT_POLL_MS))
        s.session_only = bool(d.get("session_only", False))
        s.reverse_lines_copy = bool(d.get("reverse_lines_copy", False))
        s.theme = str(d.get("theme", "flatly"))
        s.check_updates_on_launch = bool(d.get("check_updates_on_launch", True))

        s.max_history = max(5, min(HARD_MAX_HISTORY, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        return s


# -----------------------------
# GitHub update helpers (urllib only, no extra deps)
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

    def _init_state(self):
        self.data_dir = Path(user_data_dir(APP_ID, VENDOR))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"

        self.meta_path = self.data_dir / "meta.json"
        self.snippets_path = self.data_dir / "snippets.json"

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

        hist = safe_json_load(self.history_path, [])
        if isinstance(hist, list):
            hist = [x for x in hist if isinstance(x, str)]
        else:
            hist = []

        self.history = deque(hist, maxlen=self.settings.max_history)

        # Metadata (best-effort source tracking)
        meta = safe_json_load(self.meta_path, {})
        self.meta: dict[str, dict] = meta if isinstance(meta, dict) else {}

        # Snippets / templates
        snips = safe_json_load(self.snippets_path, [])
        if isinstance(snips, list):
            snips = [x for x in snips if isinstance(x, dict)]
        else:
            snips = []
        self.snippets: list[dict] = snips

        # Search options
        self.fuzzy_search_enabled = False

        # Preview match navigation (within Preview)
        self._preview_match_spans: list[tuple[int, int]] = []
        self._preview_match_i: int = 0

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
        self._selected_item_text: str | None = None  # original (non-reversed) selected item
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
            "Check Updates": "(Button)",
            "Snippets": "(Button)",
            "Format Tools": "(Button)",
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
        # Session-only means: do not persist *history*. Other settings should still persist.
        safe_json_save(self.settings_path, asdict(self.settings))
        safe_json_save(self.favs_path, list(self.favorites))
        safe_json_save(self.meta_path, self.meta)
        safe_json_save(self.snippets_path, self.snippets)
        if not self.settings.session_only:
            safe_json_save(self.history_path, list(self.history))

    # -----------------------------
    # Favorites & capacity management
    # -----------------------------
    def _ensure_favorites_present(self):
        """
        If a favorite exists but isn't in history anymore, we do not force-add it automatically
        (it may have been manually removed). This function is left as a hook if you want that behavior.
        """
        return

    # -----------------------------
    # Metadata & template helpers
    # -----------------------------
    def _get_active_window_info(self) -> tuple[str, str]:
        """Best-effort active window process name + window title (Windows only)."""
        try:
            if os.name != "nt":
                return ("", "")

            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            hwnd = user32.GetForegroundWindow()
            if not hwnd:
                return ("", "")

            # Window title
            length = user32.GetWindowTextLengthW(hwnd)
            title = ""
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value or ""

            # PID
            pid = ctypes.c_ulong(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            pid_val = int(pid.value)
            if pid_val <= 0:
                return ("", title)

            # Process name (QueryFullProcessImageNameW)
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid_val)
            if not hproc:
                return ("", title)
            try:
                size = ctypes.c_ulong(260)
                bufp = ctypes.create_unicode_buffer(260)
                if kernel32.QueryFullProcessImageNameW(hproc, 0, bufp, ctypes.byref(size)):
                    path = bufp.value
                    exe = os.path.basename(path)
                    return (exe, title)
            finally:
                kernel32.CloseHandle(hproc)

            return ("", title)
        except Exception:
            return ("", "")

    def _touch_meta(self, text: str, is_new: bool) -> None:
        if not text:
            return
        info = self.meta.get(text) if isinstance(self.meta, dict) else None
        if not isinstance(info, dict):
            info = {}

        app, title = self._get_active_window_info()
        ts = now_ts()

        if is_new and not info.get("created_at"):
            info["created_at"] = ts

        info["updated_at"] = ts
        info["copy_count"] = int(info.get("copy_count", 0)) + 1
        if app:
            info["source_app"] = app
        if title:
            info["window_title"] = title

        self.meta[text] = info

    def _move_meta_key(self, old_text: str, new_text: str) -> None:
        if not old_text or not new_text or old_text == new_text:
            return
        if old_text in self.meta and new_text not in self.meta:
            self.meta[new_text] = self.meta.get(old_text, {})
        if old_text in self.meta:
            try:
                del self.meta[old_text]
            except Exception:
                pass

    def _meta_summary(self, text: str) -> str:
        info = self.meta.get(text) if isinstance(self.meta, dict) else None
        if not isinstance(info, dict):
            return ""
        created = str(info.get("created_at") or "")
        app = str(info.get("source_app") or "")
        title = str(info.get("window_title") or "")
        cnt = str(info.get("copy_count") or "")

        parts = []
        if created:
            parts.append(f"Captured: {created}")
        if app or title:
            parts.append(f"From: {app}{' — ' if app and title else ''}{title}")
        if cnt:
            parts.append(f"Seen: {cnt}x")
        return "   |   ".join(parts)

    def _render_template(self, s: str) -> str:
        """Render a snippet/template with basic placeholders.

        Supported:
          {date}, {time}, {datetime}, {clipboard}
        """
        try:
            now = datetime.now()
            out = s
            out = out.replace("{date}", now.strftime("%Y-%m-%d"))
            out = out.replace("{time}", now.strftime("%H:%M"))
            out = out.replace("{datetime}", now.strftime("%Y-%m-%d %H:%M"))
            if "{clipboard}" in out:
                try:
                    out = out.replace("{clipboard}", str(pyperclip.paste() or ""))
                except Exception:
                    out = out.replace("{clipboard}", "")
            return out
        except Exception:
            return s


    def _prune_preserving_favorites(self, items: list[str], capacity: int) -> tuple[list[str], bool]:
        """
        Prune oldest NON-favorites first to fit capacity.
        Returns: (new_items, success)
        """
        if len(items) <= capacity:
            return items, True

        favset = set(self.favorites)

        # Remove from the front (oldest) while over capacity, skipping favorites
        i = 0
        out = items[:]
        while len(out) > capacity and i < len(out):
            if out[i] in favset:
                i += 1
                continue
            out.pop(i)

        if len(out) <= capacity:
            return out, True

        # If we are still over capacity, it means favorites alone exceed capacity
        return out, False

    def _notify_limit_reached(self):
        if self._warned_limit_reached:
            return
        self._warned_limit_reached = True
        try:
            messagebox.showinfo(
                APP_NAME,
                f"You have reached your maximum stored items ({self.settings.max_history}).\n\n"
                "Consider increasing Max history in Settings, or Clean to remove non-favorites.",
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
                "Cannot add new clipboard item because your Favorites occupy the entire capacity.\n\n"
                "Increase Max history in Settings or remove some items from Favorites.",
            )
        except Exception:
            pass

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

        existed = (text in items)

        # Stable de-dupe: remove previous occurrences
        if existed:
            items = [x for x in items if x != text]
        items.append(text)

        # Metadata update (best-effort)
        self._touch_meta(text, is_new=(not existed))


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

        # Reset selection tracking
        self._prev_sel_set = set()
        self._sel_order = []

        if select_last and self.view_items:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(tk.END)
            self.listbox.see(tk.END)
            self._on_select()

        self.status_var.set(
            f"v{APP_VERSION}   Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}"
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

        if hasattr(self, 'meta_var') and self.meta_var is not None:
            try:
                self.meta_var.set("")
            except Exception:
                pass

        self._selected_item_text = t
        display = self._get_preview_display_text_for_item(t)
        self._set_preview_text(display, mark_clean=True)

        # Metadata banner
        if hasattr(self, 'meta_var') and self.meta_var is not None:
            try:
                self.meta_var.set(self._meta_summary(t))
            except Exception:
                pass

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

        # move metadata to new key
        self._move_meta_key(old, new)

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
        self.history = deque(items, maxlen=self.settings.max_history)

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]

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
        if not messagebox.askyesno(APP_NAME, "Clean history and keep Favorites only?"):
            return

        favset = set(self.favorites)
        kept = [x for x in self.history if x in favset]

        self.history = deque(kept, maxlen=self.settings.max_history)

        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._refresh_list(select_last=True)
        self._persist()
        self.status_var.set(f"Cleaned (kept favorites) — {now_ts()}")

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
    # Format tools (Preview)
    # -----------------------------
    def _get_preview_text(self) -> str:
        return self.preview.get("1.0", tk.END).rstrip("\n")

    def _set_preview_text_dirty(self, new_text: str):
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", new_text)
        self._preview_dirty = True
        self._update_preview_dirty_ui()
        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _fmt_trim(self):
        t = self._get_preview_text()
        lines = [ln.rstrip() for ln in t.splitlines()]
        out = "\n".join(lines).strip()
        self._set_preview_text_dirty(out)

    def _fmt_normalize_newlines(self):
        t = self._get_preview_text()
        out = t.replace("\r\n", "\n").replace("\r", "\n")
        self._set_preview_text_dirty(out)

    def _fmt_remove_dupe_spaces(self):
        t = self._get_preview_text()
        out_lines = []
        for ln in t.splitlines():
            # preserve leading indentation
            m = re.match(r"^(\s*)(.*)$", ln)
            indent = m.group(1) if m else ""
            rest = m.group(2) if m else ln
            rest = re.sub(r"[ \t]{2,}", " ", rest)
            out_lines.append(indent + rest)
        self._set_preview_text_dirty("\n".join(out_lines))

    def _fmt_upper(self):
        self._set_preview_text_dirty(self._get_preview_text().upper())

    def _fmt_lower(self):
        self._set_preview_text_dirty(self._get_preview_text().lower())

    def _fmt_title(self):
        self._set_preview_text_dirty(self._get_preview_text().title())

    def _fmt_sentence(self):
        t = self._get_preview_text().strip()
        if not t:
            return
        out = t[:1].upper() + t[1:]
        self._set_preview_text_dirty(out)

    def _fmt_sanitize(self):
        t = self._get_preview_text()
        # Remove common zero-width + most control chars, but keep tab/newline
        t = re.sub(r"[\u200B-\u200D\uFEFF]", "", t)
        out = []
        for ch in t:
            o = ord(ch)
            if ch in ("\n", "\t"):
                out.append(ch)
                continue
            if o < 32:
                continue
            out.append(ch)
        self._set_preview_text_dirty("".join(out))

    def _fmt_strip_blank_lines(self):
        t = self._get_preview_text()
        lines = [ln for ln in t.splitlines() if ln.strip() != ""]
        self._set_preview_text_dirty("\n".join(lines))

    def _open_format_menu(self):
        try:
            m = tk.Menu(self, tearoff=0)
            m.add_command(label='Trim whitespace', command=self._fmt_trim)
            m.add_command(label='Normalize newlines', command=self._fmt_normalize_newlines)
            m.add_command(label='Remove duplicate spaces', command=self._fmt_remove_dupe_spaces)
            m.add_separator()
            m.add_command(label='UPPERCASE', command=self._fmt_upper)
            m.add_command(label='lowercase', command=self._fmt_lower)
            m.add_command(label='Title Case', command=self._fmt_title)
            m.add_command(label='Sentence case', command=self._fmt_sentence)
            m.add_separator()
            m.add_command(label='Sanitize (remove zero-width/control chars)', command=self._fmt_sanitize)
            m.add_command(label='Remove blank lines', command=self._fmt_strip_blank_lines)

            # Popup near mouse
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            m.tk_popup(x, y)
        finally:
            try:
                m.grab_release()
            except Exception:
                pass

    # -----------------------------
    # Actions
    # -----------------------------
    def _looks_like_url(self, s: str) -> bool:
        if not s:
            return False
        s = s.strip()
        if re.match(r"^https?://", s, re.IGNORECASE):
            return True
        # domain.tld style
        if re.match(r"^[a-z0-9.-]+\.[a-z]{2,}(/.*)?$", s, re.IGNORECASE):
            return True
        return False

    def _action_open_url(self):
        t = self._get_selected_text() or self._get_preview_text()
        if not t:
            return
        t = t.strip()
        if not self._looks_like_url(t):
            messagebox.showinfo(APP_NAME, "Selected content does not look like a URL.")
            return
        if not re.match(r"^https?://", t, re.IGNORECASE):
            t = "https://" + t
        try:
            webbrowser.open(t)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to open URL:\n{e}")

    def _action_save_selected_as_txt(self):
        t = self._get_selected_text() or self._get_preview_text()
        if not t:
            return
        path = filedialog.asksaveasfilename(
            title="Save item as",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All files", "*.*")],
            initialfile="copy2_item.txt",
        )
        if not path:
            return
        try:
            Path(path).write_text(t, encoding="utf-8")
            self.status_var.set(f"Saved to {path} — {now_ts()}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Save failed:\n{e}")

    def _action_open_in_editor(self):
        t = self._get_selected_text() or self._get_preview_text()
        if not t:
            return
        try:
            tmp = Path(tempfile.mkdtemp(prefix="copy2_open_")) / "copy2_item.txt"
            tmp.write_text(t, encoding="utf-8")
            if os.name == "nt":
                os.startfile(str(tmp))
            else:
                webbrowser.open(tmp.as_uri())
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Open failed:\n{e}")

    def _action_show_details(self):
        t = self._get_selected_text()
        if not t:
            return
        info = self.meta.get(t, {}) if isinstance(self.meta, dict) else {}
        lines = [f"Text length: {len(t)}"]
        if isinstance(info, dict):
            for k in ("created_at", "updated_at", "copy_count", "source_app", "window_title"):
                v = info.get(k)
                if v:
                    lines.append(f"{k}: {v}")
        messagebox.showinfo(APP_NAME, "\n".join(lines))

    # -----------------------------
    # Snippets / Templates
    # -----------------------------
    def _open_snippets(self):
        dlg = tk.Toplevel(self)
        dlg.title("Snippets")
        dlg.resizable(True, True)
        dlg.transient(self)
        dlg.grab_set()

        if USE_TTKB:
            Frame = tb.Frame
            Label = tb.Label
            Button = tb.Button
            Entry = tb.Entry
        else:
            from tkinter import ttk as _ttk  # type: ignore
            Frame = _ttk.Frame
            Label = _ttk.Label
            Button = _ttk.Button
            Entry = _ttk.Entry

        root = Frame(dlg, padding=12)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        Label(root, text="Snippets (templates)", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")

        lb = tk.Listbox(root, activestyle="dotbox", exportselection=False)
        lb.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        sb = tk.Scrollbar(root, orient="vertical", command=lb.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=(10, 0))
        lb.configure(yscrollcommand=sb.set)

        hint = ("Placeholders: {date}, {time}, {datetime}, {clipboard}")
        Label(root, text=hint, foreground="#666").grid(row=2, column=0, sticky="w", pady=(10, 0))

        def refresh():
            lb.delete(0, tk.END)
            for s in self.snippets:
                name = str(s.get('name') or '')
                trig = str(s.get('trigger') or '')
                lb.insert(tk.END, f"{trig} — {name}" if name else trig)

        def get_sel_index() -> int | None:
            try:
                sel = lb.curselection()
                if not sel:
                    return None
                return int(sel[0])
            except Exception:
                return None

        def add_or_edit(idx: int | None):
            data = self.snippets[idx] if idx is not None and 0 <= idx < len(self.snippets) else {'name': '', 'trigger': '', 'body': ''}

            ed = tk.Toplevel(dlg)
            ed.title('Edit Snippet' if idx is not None else 'Add Snippet')
            ed.resizable(True, True)
            ed.transient(dlg)
            ed.grab_set()

            frm = Frame(ed, padding=12)
            frm.pack(fill=tk.BOTH, expand=True)
            frm.columnconfigure(1, weight=1)
            frm.rowconfigure(2, weight=1)

            Label(frm, text='Name').grid(row=0, column=0, sticky='w')
            name_var = tk.StringVar(value=str(data.get('name') or ''))
            Entry(frm, textvariable=name_var).grid(row=0, column=1, sticky='ew', padx=(10, 0))

            Label(frm, text='Trigger').grid(row=1, column=0, sticky='w', pady=(10, 0))
            trig_var = tk.StringVar(value=str(data.get('trigger') or ''))
            Entry(frm, textvariable=trig_var).grid(row=1, column=1, sticky='ew', padx=(10, 0), pady=(10, 0))

            Label(frm, text='Body').grid(row=2, column=0, sticky='nw', pady=(10, 0))
            body = tk.Text(frm, wrap='word', undo=True, height=10)
            body.grid(row=2, column=1, sticky='nsew', padx=(10, 0), pady=(10, 0))
            body.insert('1.0', str(data.get('body') or ''))

            btns = Frame(frm)
            btns.grid(row=3, column=0, columnspan=2, sticky='e', pady=(12, 0))

            def save_snip():
                trig = trig_var.get().strip()
                if not trig:
                    messagebox.showwarning(APP_NAME, 'Trigger is required.')
                    return
                new_data = {
                    'name': name_var.get().strip(),
                    'trigger': trig,
                    'body': body.get('1.0', tk.END).rstrip('\n'),
                }

                # Enforce unique triggers
                for i, s in enumerate(self.snippets):
                    if i != (idx if idx is not None else -1) and str(s.get('trigger') or '').strip() == trig:
                        messagebox.showwarning(APP_NAME, 'That trigger already exists.')
                        return

                if idx is None:
                    self.snippets.append(new_data)
                else:
                    self.snippets[idx] = new_data

                self._persist()
                refresh()
                ed.destroy()

            Button(btns, text='Cancel', command=ed.destroy).grid(row=0, column=0, padx=(0, 10))
            Button(btns, text='Save', command=save_snip).grid(row=0, column=1)

        def delete_sel():
            idx = get_sel_index()
            if idx is None:
                return
            if not messagebox.askyesno(APP_NAME, 'Delete this snippet?'):
                return
            try:
                self.snippets.pop(idx)
            except Exception:
                return
            self._persist()
            refresh()

        def copy_sel():
            idx = get_sel_index()
            if idx is None:
                return
            s = self.snippets[idx]
            body = self._render_template(str(s.get('body') or ''))
            if not body:
                return
            try:
                pyperclip.copy(body)
            except Exception:
                return
            self._add_history_item(body)
            self.status_var.set(f"Snippet copied — {now_ts()}")

        def insert_into_preview():
            idx = get_sel_index()
            if idx is None:
                return
            s = self.snippets[idx]
            body = self._render_template(str(s.get('body') or ''))
            if not body:
                return
            try:
                self.preview.insert(tk.INSERT, body)
                self._mark_preview_dirty()
            except Exception:
                pass

        actions = Frame(root)
        actions.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(12, 0))

        Button(actions, text='Add', command=lambda: add_or_edit(None)).pack(side=tk.LEFT)
        Button(actions, text='Edit', command=lambda: add_or_edit(get_sel_index())).pack(side=tk.LEFT, padx=(8, 0))
        Button(actions, text='Delete', command=delete_sel).pack(side=tk.LEFT, padx=(8, 0))

        Button(actions, text='Copy to Clipboard', command=copy_sel).pack(side=tk.RIGHT)
        Button(actions, text='Insert into Preview', command=insert_into_preview).pack(side=tk.RIGHT, padx=(0, 8))

        refresh()


    # -----------------------------
    # Find / Highlight
    # -----------------------------
    def _highlight_query_in_preview(self, query: str):
        self.preview.tag_remove("match", "1.0", tk.END)
        self._preview_match_spans = []
        self._preview_match_i = 0
        if not query:
            return

        text = self.preview.get("1.0", tk.END)
        if not text.strip():
            return

        pattern = re.compile(re.escape(query), re.IGNORECASE)
        for m in pattern.finditer(text):
            self._preview_match_spans.append((m.start(), m.end()))

        if not self._preview_match_spans:
            return

        try:
            self.preview.tag_config("match", background="#2b78ff", foreground="white")
        except Exception:
            pass

        for a, b in self._preview_match_spans:
            start_index = f"1.0+{a}c"
            end_index = f"1.0+{b}c"
            self.preview.tag_add("match", start_index, end_index)

        # Scroll to first match
        a0, _b0 = self._preview_match_spans[0]
        self.preview.see(f"1.0+{a0}c")

    def _search(self):
        q = self.search_var.get().strip()
        self.search_query = q
        self.search_matches = []
        self.search_index = 0

        if not q:
            self.status_var.set("Search cleared.")
            self.preview.tag_remove("match", "1.0", tk.END)
            return

        ql = q.lower()
        items = list(self.history)

        # Fuzzy search (optional)
        fuzzy = False
        if hasattr(self, 'fuzzy_var') and self.fuzzy_var is not None:
            try:
                fuzzy = bool(self.fuzzy_var.get())
            except Exception:
                fuzzy = False

        if not fuzzy:
            for item in items:
                if ql in item.lower():
                    self.search_matches.append(item)
        else:
            scored: list[tuple[float, str]] = []
            for item in items:
                il = item.lower()
                if ql in il:
                    # strong signal
                    scored.append((2.0, item))
                    continue
                # Sequence similarity
                r = difflib.SequenceMatcher(None, ql, il[: min(len(il), 600)]).ratio()
                if r >= 0.35:
                    scored.append((r, item))
            scored.sort(key=lambda x: x[0], reverse=True)
            self.search_matches = [it for _s, it in scored[:250]]

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

    def _jump_preview_match(self, direction: int):
        if not self._preview_match_spans:
            return
        self._preview_match_i = (self._preview_match_i + direction) % len(self._preview_match_spans)
        a, _b = self._preview_match_spans[self._preview_match_i]
        self.preview.see(f"1.0+{a}c")


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
            Separator = tb.Separator
        else:
            from tkinter import ttk as _ttk  # type: ignore
            Frame = _ttk.Frame
            Label = _ttk.Label
            Entry = _ttk.Entry
            Button = _ttk.Button
            Checkbutton = _ttk.Checkbutton
            Combobox = _ttk.Combobox
            Separator = _ttk.Separator

        frm = Frame(dlg, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        max_var = tk.StringVar(value=str(self.settings.max_history))
        poll_var = tk.StringVar(value=str(self.settings.poll_ms))
        sess_var = tk.BooleanVar(value=self.settings.session_only)
        upd_var = tk.BooleanVar(value=self.settings.check_updates_on_launch)

        theme_var = tk.StringVar(value=self.settings.theme)
        themes = ["flatly", "litera", "cosmo", "sandstone", "minty", "darkly", "superhero", "cyborg"]

        Label(frm, text=f"Max history (5–{HARD_MAX_HISTORY}):").grid(row=0, column=0, sticky="w")
        ent_mh = Entry(frm, textvariable=max_var, width=12)
        ent_mh.grid(row=0, column=1, sticky="w", padx=(10, 0))

        Label(frm, text="Poll interval ms (100–5000):").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ent_pm = Entry(frm, textvariable=poll_var, width=12)
        ent_pm.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(10, 0))

        Checkbutton(frm, text="Session-only (do not save history)", variable=sess_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

        Checkbutton(frm, text="Check updates on launch", variable=upd_var).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

        if USE_TTKB:
            Label(frm, text="Theme:").grid(row=4, column=0, sticky="w", pady=(10, 0))
            cb = Combobox(frm, textvariable=theme_var, values=themes, width=18, state="readonly")
            cb.grid(row=4, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        else:
            Label(frm, text="(Install ttkbootstrap for modern themes)", foreground="#666").grid(
                row=4, column=0, columnspan=2, sticky="w", pady=(10, 0)
            )

        Separator(frm).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(14, 10))

        # Controls section
        Label(frm, text="Controls / Hotkeys", font=("Segoe UI", 10, "bold")).grid(row=6, column=0, columnspan=2, sticky="w")

        hot_txt = "\n".join([f"{k}: {v}" for k, v in self.HOTKEYS.items()])
        lbl_hot = Label(frm, text=hot_txt, justify="left")
        lbl_hot.grid(row=7, column=0, columnspan=2, sticky="w")

        btns = Frame(frm)
        btns.grid(row=8, column=0, columnspan=2, sticky="e", pady=(14, 0))

        def save():
            try:
                mh = int(max_var.get().strip())
                pm = int(poll_var.get().strip())
            except Exception:
                messagebox.showerror(APP_NAME, "Please enter valid integers.")
                return

            mh = max(5, min(HARD_MAX_HISTORY, mh))
            pm = max(100, min(5000, pm))

            # If hard cap is reached, warn clearly
            if mh >= HARD_MAX_HISTORY:
                self._notify_hard_cap_reached()

            self.settings.max_history = mh
            self.settings.poll_ms = pm
            self.settings.session_only = bool(sess_var.get())
            self.settings.check_updates_on_launch = bool(upd_var.get())

            if USE_TTKB:
                self.settings.theme = str(theme_var.get()).strip() or "flatly"

            # Prune existing history to new capacity (preserve favorites)
            items = list(self.history)
            pruned, ok = self._prune_preserving_favorites(items, mh)
            if not ok:
                self._notify_favorites_blocking()
                # If favorites exceed cap, keep as many as possible but never delete favorites automatically.
                # We just keep current list and tell user to increase cap or reduce favorites.
            self.history = deque(pruned[-mh:], maxlen=mh)

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
        self.menu.add_command(label="Copy", command=self._copy_selected)
        self.menu.add_command(label="Open as URL", command=self._action_open_url)
        self.menu.add_command(label="Open in Editor", command=self._action_open_in_editor)
        self.menu.add_command(label="Save as .txt...", command=self._action_save_selected_as_txt)
        self.menu.add_command(label="Details...", command=self._action_show_details)
        self.menu.add_command(label="Favorite / Unfavorite", command=self._toggle_favorite_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Save Preview Edit", command=self._save_preview_edits)
        self.menu.add_command(label="Revert Preview Edit", command=self._revert_preview_edits)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self._delete_selected)
        self.menu.add_command(label="Combine Selected", command=self._combine_selected)

        self.menu.add_separator()
        self.menu.add_command(label="Snippets...", command=self._open_snippets)

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

            self.fuzzy_var = tk.BooleanVar(value=False)
            fuzzy_btn = tb.Checkbutton(search_box, text="Fuzzy", variable=self.fuzzy_var, bootstyle="round-toggle")
            fuzzy_btn.pack(side=LEFT, padx=(10, 0))
            ToolTip(fuzzy_btn, "Fuzzy search (approximate matching)")

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

            btn_snips = tb.Button(action_box, text="Snippets", command=self._open_snippets, bootstyle=SECONDARY)
            btn_snips.pack(side=LEFT, padx=(0, 8))
            ToolTip(btn_snips, "Manage snippets / templates")

            btn_settings = tb.Button(action_box, text="Settings", command=self._open_settings, bootstyle=SECONDARY)
            btn_settings.pack(side=LEFT)
            ToolTip(btn_settings, "Settings (includes hotkeys list)")

            tb.Separator(root).pack(fill=X, pady=(0, 10))

            # Body split
            body = tb.Frame(root)
            body.pack(fill=BOTH, expand=True)
            body.columnconfigure(0, weight=1)
            body.columnconfigure(1, weight=2)
            body.rowconfigure(0, weight=1)

            # Left pane (History)
            left = tb.Labelframe(body, text="History", padding=10)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            left.rowconfigure(2, weight=1)
            left.columnconfigure(0, weight=1)

            filters = tb.Frame(left)
            filters.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            filters.columnconfigure(3, weight=1)

            self.filter_var = tk.StringVar(value="all")
            tb.Radiobutton(filters, text="All", value="all", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=0, padx=(0, 8))
            tb.Radiobutton(filters, text="Favorites", value="fav", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=1, padx=(0, 8))

            btn_clean = tb.Button(filters, text="Clean", command=self._clean_keep_favorites, bootstyle=DANGER)
            btn_clean.grid(row=0, column=4, sticky="e")
            ToolTip(btn_clean, "Clean (keeps favorites)  (Ctrl+L)")

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

            # Right pane (Preview)
            right = tb.Labelframe(body, text="Preview (Editable)", padding=10)
            right.grid(row=0, column=1, sticky="nsew")
            right.rowconfigure(1, weight=1)
            right.columnconfigure(0, weight=1)

            preview_actions = tb.Frame(right)
            preview_actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            preview_actions.columnconfigure(0, weight=1)

            self.reverse_var = tk.BooleanVar(value=self.settings.reverse_lines_copy)
            rev_btn = tb.Checkbutton(
                preview_actions,
                text="Reverse-lines copy",
                variable=self.reverse_var,
                command=self._toggle_reverse_lines,
                bootstyle="round-toggle",
            )
            rev_btn.pack(side=LEFT)
            ToolTip(rev_btn, "Reverse lines in Preview and Copy")

            # Metadata banner
            self.meta_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.meta_var, font=("Segoe UI", 9), foreground="#666").pack(side=LEFT, padx=(12, 0))

            self.preview_dirty_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.preview_dirty_var).pack(side=LEFT, padx=(10, 0))

            # Format tools menu
            fmt_btn = tb.Button(preview_actions, text="Format", bootstyle=OUTLINE, command=self._open_format_menu)
            fmt_btn.pack(side=RIGHT, padx=(0, 8))
            ToolTip(fmt_btn, "Format tools (trim, case, sanitize, etc.)")

            self.revert_btn = tb.Button(preview_actions, text="Revert", command=self._revert_preview_edits, bootstyle=WARNING)
            self.revert_btn.pack(side=RIGHT, padx=(0, 8))
            ToolTip(self.revert_btn, "Revert (Esc)")

            self.save_btn = tb.Button(preview_actions, text="Save Edit", command=self._save_preview_edits, bootstyle=SUCCESS)
            self.save_btn.pack(side=RIGHT, padx=(0, 8))
            ToolTip(self.save_btn, "Save (Ctrl+S)")

            btn_copy_prev = tb.Button(preview_actions, text="Copy Preview", command=self._copy_selected, bootstyle=SUCCESS)
            btn_copy_prev.pack(side=RIGHT)
            ToolTip(btn_copy_prev, "Copy (Ctrl+C)")

            self.preview = tk.Text(right, wrap="word", undo=True)
            self.preview.grid(row=1, column=0, sticky="nsew")
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            sb2 = tb.Scrollbar(right, orient="vertical", command=self.preview.yview)
            sb2.grid(row=1, column=1, sticky="ns")
            self.preview.configure(yscrollcommand=sb2.set)

            # Bottom status bar
            bottom = tb.Frame(root)
            bottom.pack(fill=X, pady=(10, 0))
            self.status_var = tk.StringVar(value=f"v{APP_VERSION}   Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}")
            tb.Label(bottom, textvariable=self.status_var).pack(side=LEFT)

            self._update_preview_dirty_ui()

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
            self.bind("<Control-Shift-Next>", lambda e: (self._jump_preview_match(1), "break"))
            self.bind("<Control-Shift-Prior>", lambda e: (self._jump_preview_match(-1), "break"))
            self.bind("<Control-Alt-s>", lambda e: (self._open_snippets(), "break"))

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
            self.reverse_var = tk.BooleanVar(value=self.settings.reverse_lines_copy)
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
