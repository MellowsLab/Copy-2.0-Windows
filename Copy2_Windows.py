"""
Copy 2.0 (Windows) — Modern UI Edition

1) Favorites are protected
   - “Clean” removes ONLY non-favorites; favorites remain.
   - Favorites are NEVER auto-evicted when max_history is reached.
   - If max_history is reached and ONLY favorites remain (or favorites >= max_history), new items will NOT be added;
     you will be prompted to increase max_history.
2) Soft cap + hard cap notifications
   - Soft cap: when you hit your configured max_history, you get a prompt (once per run) suggesting increasing it.
   - Hard cap: when you hit hard-coded storage limits, you get “allocated memory used up” style prompt; no new items are stored.
3) Controls / Hotkeys viewer
   - Settings dialog includes a “Controls” section showing all hotkeys.
   - Main buttons include hover tooltips showing hotkeys (ttkbootstrap) or a simple fallback tooltip.
4) Import/Export hardened
   - Export includes history + favorites + settings.
   - Import merges, stable de-dupes, ensures favorites exist in history, then enforces caps safely.
5) Auto-updater (GitHub Releases)
   - On launch: checks GitHub releases for a newer tag/version.
   - “Check Updates” button to manually check.
   - If update available: prompt to update now.
   - Downloads the Release Asset (preferred) or falls back to a .zip link inside release notes (user-attachments).
   - Validates download (size/signature) and uses safe swap (backup + rollback) then restarts.
   - Logs:
       <data_dir>/update_check.log
       <data_dir>/update_install.log

Dependencies:
  pyperclip
  platformdirs
Recommended (modern UI):
  ttkbootstrap

Build:
  python -m PyInstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py
"""

from __future__ import annotations

import json
import os
import re
import ssl
import sys
import time
import shutil
import zipfile
import tempfile
import threading
import subprocess
import urllib.request
import urllib.error
import webbrowser
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
    try:
        from ttkbootstrap.tooltip import ToolTip as TBToolTip
    except Exception:
        TBToolTip = None
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
# App / Update config
# -----------------------------
APP_NAME = "Copy 2.0"
APP_ID = "copy2"
VENDOR = "MellowsLab"

# IMPORTANT: keep this in sync with your release tags, e.g. "v1.0.2"
APP_VERSION = "v1.0.3"

# GitHub repo to check
GITHUB_OWNER = "MellowsLab"
GITHUB_REPO = "Copy-2.0-Windows"

# Optional: hint to pick the right asset if multiple ZIPs exist
GITHUB_ASSET_NAME_HINT = "Copy2"  # substring match

# Update behavior
CHECK_UPDATES_ON_STARTUP = True
STARTUP_UPDATE_DELAY_MS = 900

DEFAULT_MAX_HISTORY = 50
DEFAULT_POLL_MS = 400

# -----------------------------
# Hard-coded storage limits (“allocated memory”)
# -----------------------------
HARD_MAX_ITEMS = 2000          # total items in history (favorites included)
HARD_MAX_TOTAL_CHARS = 5_000_000  # sum of string lengths across all items (approx memory bound)

# PyInstaller onefile sanity: minimum expected EXE size
MIN_EXE_BYTES = 5_000_000


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


def parse_version_tag(tag: str) -> tuple[int, int, int, int]:
    """
    Accepts: v1.2.3, 1.2.3, 1.2, v1.2.3-beta (beta -> lower priority)
    Returns (major, minor, patch, preflag) where preflag=1 means prerelease.
    """
    t = (tag or "").strip()
    if t.lower().startswith("v"):
        t = t[1:]
    preflag = 0
    # crude prerelease detection
    if re.search(r"[a-zA-Z]", t):
        preflag = 1
    # keep numeric parts only
    m = re.match(r"^\s*(\d+)(?:\.(\d+))?(?:\.(\d+))?", t)
    if not m:
        return (0, 0, 0, 1)
    maj = int(m.group(1) or 0)
    minr = int(m.group(2) or 0)
    pat = int(m.group(3) or 0)
    return (maj, minr, pat, preflag)


def is_newer(latest: str, installed: str) -> bool:
    return parse_version_tag(latest) > parse_version_tag(installed)


class SimpleTooltip:
    """Small tooltip fallback if ttkbootstrap ToolTip is unavailable."""
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _e=None):
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 10
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
            self.tip = tk.Toplevel(self.widget)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{x}+{y}")
            lbl = tk.Label(self.tip, text=self.text, background="#111", foreground="#fff",
                           relief="solid", borderwidth=1, padx=6, pady=3)
            lbl.pack()
        except Exception:
            self.tip = None

    def _hide(self, _e=None):
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

        # soft bounds for user config
        s.max_history = max(5, min(500, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        return s


class Copy2AppBase:
    """Shared logic for both ttkbootstrap and ttk fallback variants."""

    # -----------------------------
    # init / persistence
    # -----------------------------
    def _init_state(self):
        self.data_dir = Path(user_data_dir(APP_ID, VENDOR))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"

        self.settings = Settings.from_dict(safe_json_load(self.settings_path, {}))

        # Store history as a plain list (newest last). We enforce caps ourselves to protect favorites.
        raw_hist = safe_json_load(self.history_path, [])
        self.history: list[str] = [x for x in raw_hist if isinstance(x, str)]

        raw_favs = safe_json_load(self.favs_path, [])
        self.favorites: list[str] = [x for x in raw_favs if isinstance(x, str)]

        # Ensure favorites exist in history
        for f in self.favorites:
            if f not in self.history:
                self.history.append(f)

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

        # Cap notifications (per run)
        self._notified_soft_cap = False
        self._notified_hard_cap = False

        # Update check error stash
        self._last_update_error = ""

        # Ensure caps are enforced at startup
        self._enforce_caps_and_persist()

    def _persist(self):
        if self.settings.session_only:
            return
        safe_json_save(self.settings_path, asdict(self.settings))
        safe_json_save(self.history_path, list(self.history))
        safe_json_save(self.favs_path, list(self.favorites))

    def _log_update_check(self, msg: str):
        try:
            p = self.data_dir / "update_check.log"
            p.parent.mkdir(parents=True, exist_ok=True)
            with p.open("a", encoding="utf-8", errors="ignore") as f:
                f.write(f"{now_ts()}  {msg}\n")
        except Exception:
            pass

    def _log_update_install(self, msg: str):
        try:
            p = self.data_dir / "update_install.log"
            p.parent.mkdir(parents=True, exist_ok=True)
            with p.open("a", encoding="utf-8", errors="ignore") as f:
                f.write(f"{now_ts()}  {msg}\n")
        except Exception:
            pass

    # -----------------------------
    # Caps enforcement (favorites protected)
    # -----------------------------
    def _total_chars(self) -> int:
        return sum(len(x) for x in self.history)

    def _hard_cap_reached(self) -> bool:
        if len(self.history) >= HARD_MAX_ITEMS:
            return True
        if self._total_chars() >= HARD_MAX_TOTAL_CHARS:
            return True
        return False

    def _enforce_caps_and_persist(self):
        """
        Enforce soft cap (settings.max_history) by evicting oldest NON-favorites first.
        If favorites prevent eviction enough, we allow the list to remain above max_history.
        """
        # Stable de-dupe: keep latest occurrence
        seen = set()
        out = []
        for item in reversed(self.history):
            if item not in seen:
                out.append(item)
                seen.add(item)
        out.reverse()
        self.history = out

        # Ensure favorites exist in history
        for f in self.favorites:
            if f not in self.history:
                self.history.append(f)

        # Enforce hard caps by evicting non-favorites oldest-first; if impossible, keep and block new writes later
        # NOTE: we never delete favorites automatically.
        def evict_one_nonfav() -> bool:
            for i, item in enumerate(self.history):
                if item not in self.favorites:
                    del self.history[i]
                    return True
            return False

        while len(self.history) > HARD_MAX_ITEMS:
            if not evict_one_nonfav():
                break
        while self._total_chars() > HARD_MAX_TOTAL_CHARS:
            if not evict_one_nonfav():
                break

        # Soft cap: try to get down to settings.max_history by removing non-favorites
        while len(self.history) > self.settings.max_history:
            if not evict_one_nonfav():
                # can't evict without removing favorites
                break

        self._persist()

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

    def _notify_soft_cap_once(self):
        if self._notified_soft_cap:
            return
        self._notified_soft_cap = True
        try:
            if messagebox.askyesno(
                APP_NAME,
                f"You have reached your maximum stored items ({self.settings.max_history}).\n\n"
                "Older non-favorite items will be removed to make space.\n"
                "Favorites are protected.\n\n"
                "Would you like to open Settings to increase the limit?"
            ):
                self._open_settings()
        except Exception:
            pass

    def _notify_hard_cap_once(self):
        if self._notified_hard_cap:
            return
        self._notified_hard_cap = True
        try:
            messagebox.showwarning(
                APP_NAME,
                "You have used up all allocated storage for clipboard history.\n\n"
                "No new items will be stored until you remove old items (preferably non-favorites),\n"
                "or reduce total stored content.\n\n"
                "Favorites are protected from automatic removal."
            )
        except Exception:
            pass

    def _add_history_item(self, text: str):
        text = (text or "").strip("\r\n")
        if not text:
            return

        # Hard cap: refuse to add if we cannot make room without deleting favorites
        if self._hard_cap_reached():
            # Try evicting non-favorites to get under hard caps
            before = len(self.history)
            self._enforce_caps_and_persist()
            if self._hard_cap_reached() and len(self.history) == before:
                self._notify_hard_cap_once()
                return

        # De-dupe: remove existing occurrence
        if text in self.history:
            self.history = [x for x in self.history if x != text]
        self.history.append(text)

        # If we are at/over soft cap, enforce (evict non-favorites only)
        if len(self.history) >= self.settings.max_history:
            self._notify_soft_cap_once()

        # Enforce caps (soft + hard)
        self._enforce_caps_and_persist()

        # If we still exceed soft cap and cannot evict (favorites block), refuse storing NEW non-favorite items
        if len(self.history) > self.settings.max_history:
            # If text itself is not favorite and we couldn't make room, undo append
            if text not in self.favorites:
                # remove the newest we just added
                try:
                    self.history = [x for x in self.history if x != text]
                except Exception:
                    pass
                self._persist()
                try:
                    messagebox.showwarning(
                        APP_NAME,
                        f"Cannot store more items because favorites are protected and your limit is {self.settings.max_history}.\n\n"
                        "Increase Max history in Settings, or remove some favorites."
                    )
                except Exception:
                    pass
                self._refresh_list(select_last=True)
                return

        self._refresh_list(select_last=True)

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
            f"Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}"
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

        # Replace old in history
        self.history = [x for x in self.history if x != old]
        self.history.append(new)

        # Maintain favorite mapping
        if old in self.favorites:
            self.favorites = [new if x == old else x for x in self.favorites]

        self._selected_item_text = new
        self._preview_dirty = False

        self._enforce_caps_and_persist()
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

        # Remove from history
        self.history = [x for x in self.history if x != t]

        # Remove from favorites if present
        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]

        if self._selected_item_text == t:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)

        self._enforce_caps_and_persist()
        self._refresh_list(select_last=True)

    def _clean_history(self):
        """
        "Clean" keeps favorites; removes non-favorites only.
        """
        if not messagebox.askyesno(APP_NAME, "Clean history?\n\nThis will remove ALL non-favorite items.\nFavorites will remain."):
            return
        self.history = [x for x in self.history if x in self.favorites]
        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._enforce_caps_and_persist()
        self._refresh_list(select_last=True)

    def _clear_all_history(self):
        """
        Full wipe (including favorites) – kept as a context-menu option.
        """
        if not messagebox.askyesno(APP_NAME, "Clear ALL history items (including favorites)?"):
            return
        self.history = []
        self.favorites = []
        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._persist()
        self._refresh_list(select_last=True)

    def _toggle_favorite_selected(self):
        t = self._get_selected_text()
        if not t:
            return

        if t in self.favorites:
            self.favorites = [x for x in self.favorites if x != t]
        else:
            self.favorites.append(t)
            # ensure favorite exists in history
            if t not in self.history:
                self.history.append(t)

        self._enforce_caps_and_persist()
        self._refresh_list()
        self.status_var.set(f"Favorites updated — {now_ts()}")

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

        for item in self.history:
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
            "exported_at": now_ts(),
            "app": APP_NAME,
            "version": APP_VERSION,
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

            favs = list(self.favorites)
            if isinstance(favorites, list):
                for f in favorites:
                    if isinstance(f, str) and f.strip() and f not in favs:
                        favs.append(f)

            # Stable de-dupe keep latest
            seen = set()
            out = []
            for item in reversed(merged):
                if item not in seen:
                    out.append(item)
                    seen.add(item)
            out.reverse()

            self.history = out
            self.favorites = favs

            # Ensure favorites exist in history
            for f in self.favorites:
                if f not in self.history:
                    self.history.append(f)

            self._enforce_caps_and_persist()
            self._refresh_list(select_last=True)
            self.status_var.set(f"Imported — {path}")

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")

    # -----------------------------
    # Settings dialog (with Controls section)
    # -----------------------------
    def _hotkeys_text(self) -> str:
        return (
            "Hotkeys / Controls\n"
            "------------------\n"
            "Ctrl+F  : Focus Search\n"
            "Enter   : Run Search\n"
            "Ctrl+C  : Copy Preview\n"
            "Delete  : Delete Selected\n"
            "Ctrl+S  : Save Preview Edit\n"
            "Esc     : Revert Preview Edit\n"
            "Ctrl+E  : Export\n"
            "Ctrl+I  : Import\n"
            "Ctrl+L  : Clean (non-favorites only)\n"
            "\n"
            "Notes:\n"
            "- Favorites are protected from Clean and auto-eviction.\n"
            "- Combine uses Ctrl+Click to multi-select in click order.\n"
        )

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

        # Controls section
        Separator(frm).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 10))
        Label(frm, text="Controls / Hotkeys:").grid(row=5, column=0, columnspan=2, sticky="w")

        txt = tk.Text(frm, width=54, height=10, wrap="word")
        txt.grid(row=6, column=0, columnspan=2, sticky="ew")
        txt.insert("1.0", self._hotkeys_text())
        txt.configure(state="disabled")

        btns = Frame(frm)
        btns.grid(row=7, column=0, columnspan=2, sticky="e", pady=(14, 0))

        def save():
            try:
                mh = int(max_var.get().strip())
                pm = int(poll_var.get().strip())
            except Exception:
                messagebox.showerror(APP_NAME, "Please enter valid integers.")
                return

            mh = max(5, min(500, mh))
            pm = max(100, min(5000, pm))

            self.settings.max_history = mh
            self.settings.poll_ms = pm
            self.settings.session_only = bool(sess_var.get())

            if USE_TTKB:
                self.settings.theme = str(theme_var.get()).strip() or "flatly"

            self._enforce_caps_and_persist()
            self._refresh_list(select_last=True)

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
        self.menu.add_command(label="Favorite / Unfavorite", command=self._toggle_favorite_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Save Preview Edit", command=self._save_preview_edits)
        self.menu.add_command(label="Revert Preview Edit", command=self._revert_preview_edits)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self._delete_selected)
        self.menu.add_command(label="Combine Selected", command=self._combine_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Clean (keep favorites)", command=self._clean_history)
        self.menu.add_command(label="Clear ALL (including favorites)", command=self._clear_all_history)

        def popup(event):
            try:
                idx = self.listbox.nearest(event.y)
                if idx >= 0:
                    if not (event.state & 0x0004):
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
    # Updater: GitHub checks + apply
    # -----------------------------
    def _is_frozen(self) -> bool:
        return bool(getattr(sys, "frozen", False)) and hasattr(sys, "_MEIPASS")

    def _github_api_latest_url(self) -> str:
        return f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

    def _fetch_latest_release_info(self) -> dict | None:
        """
        Returns: {version, html_url, notes, asset_name, asset_url, asset_size}
        Prefers Release Assets; falls back to finding a .zip link in release notes if no assets exist.
        """
        url = self._github_api_latest_url()
        self._last_update_error = ""
        try:
            ctx = ssl.create_default_context()
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": f"{APP_NAME}/{APP_VERSION}",
                    "Accept": "application/vnd.github+json",
                },
                method="GET",
            )
            with urllib.request.urlopen(req, timeout=10, context=ctx) as r:
                status = getattr(r, "status", 200)
                raw = r.read().decode("utf-8", errors="replace")
            if status >= 400:
                self._log_update_check(f"HTTP {status}: {raw[:900]}")
                self._last_update_error = f"GitHub returned HTTP {status}."
                return None

            data = json.loads(raw)
            tag = str(data.get("tag_name", "")).strip()
            html_url = str(data.get("html_url", "")).strip()
            notes = str(data.get("body", "")).strip()

            assets = data.get("assets", [])
            if not isinstance(assets, list):
                assets = []

            zip_like = []
            exe_like = []
            hint = (GITHUB_ASSET_NAME_HINT or "").lower()

            for a in assets:
                if not isinstance(a, dict):
                    continue
                name = str(a.get("name", "")).strip()
                dl = str(a.get("browser_download_url", "")).strip()
                size = int(a.get("size", 0) or 0)
                ctype = str(a.get("content_type", "")).lower()
                if not name or not dl:
                    continue
                nlow = name.lower()
                is_zip = nlow.endswith(".zip") or ("zip" in ctype)
                is_exe = nlow.endswith(".exe") or ("msdownload" in ctype)
                entry = {"name": name, "url": dl, "size": size}
                if is_zip:
                    zip_like.append(entry)
                elif is_exe:
                    exe_like.append(entry)

            def pick_best(cands):
                if not cands:
                    return None
                if hint:
                    for c in cands:
                        if hint in c["name"].lower():
                            return c
                return sorted(cands, key=lambda x: x.get("size", 0), reverse=True)[0]

            chosen = pick_best(zip_like) or pick_best(exe_like)

            # Fallback: notes body link (user-attachments or any .zip)
            if not chosen:
                m = re.search(r"(https://github\.com/user-attachments/files/[^\s)]+\.zip)", notes, re.IGNORECASE)
                if not m:
                    m = re.search(r"(https://github\.com/[^\s)]+\.zip)", notes, re.IGNORECASE)
                if m:
                    chosen = {"name": "linked_zip_from_notes.zip", "url": m.group(1), "size": 0}

            info = {
                "version": tag,
                "html_url": html_url,
                "notes": notes,
                "asset_name": chosen["name"] if chosen else "",
                "asset_url": chosen["url"] if chosen else "",
                "asset_size": int(chosen["size"]) if chosen else 0,
            }

            self._log_update_check(f"OK tag={info['version']} asset={info['asset_name'] or '(none)'}")
            return info if info["version"] else None

        except urllib.error.HTTPError as e:
            try:
                body = e.read().decode("utf-8", errors="replace")
            except Exception:
                body = ""
            self._log_update_check(f"HTTPError {e.code}: {body[:900]}")
            self._last_update_error = f"GitHub returned HTTP {e.code}."
            return None
        except urllib.error.URLError as e:
            self._log_update_check(f"URLError: {e}")
            self._last_update_error = f"Network error: {e}"
            return None
        except Exception as e:
            self._log_update_check(f"Exception: {repr(e)}")
            self._last_update_error = f"Unexpected error: {repr(e)}"
            return None

    def _check_updates_async(self, prompt_if_new: bool = True):
        def worker():
            info = self._fetch_latest_release_info()
            def ui():
                if info is None:
                    msg = self._last_update_error or "Could not check for updates."
                    messagebox.showwarning(APP_NAME, f"{msg}\n\nLog: {self.data_dir / 'update_check.log'}")
                    self.status_var.set(f"Update check failed — {now_ts()}")
                    return

                latest = str(info.get("version", "")).strip()
                if not latest:
                    messagebox.showwarning(APP_NAME, "No version found on latest release.")
                    return

                if not is_newer(latest, APP_VERSION):
                    if prompt_if_new:
                        messagebox.showinfo(APP_NAME, f"You are up to date.\n\nInstalled: {APP_VERSION}\nLatest: {latest}")
                    self.status_var.set(f"Up to date — {now_ts()}")
                    return

                if not prompt_if_new:
                    # Still prompt on startup per your requirement
                    pass

                asset_url = str(info.get("asset_url", "")).strip()
                html_url = str(info.get("html_url", "")).strip()

                if not asset_url:
                    resp = messagebox.askyesno(
                        APP_NAME,
                        f"Update available.\n\nInstalled: {APP_VERSION}\nLatest: {latest}\n\n"
                        "No downloadable ZIP/EXE asset was found on the latest release.\n"
                        "This usually means the release asset was not uploaded under Assets.\n\n"
                        "Open the release page?"
                    )
                    if resp and html_url:
                        webbrowser.open(html_url)
                    return

                resp = messagebox.askyesno(
                    APP_NAME,
                    f"Update available.\n\nInstalled: {APP_VERSION}\nLatest: {latest}\n\n"
                    "Would you like to download and install it now?\n"
                    "The app will restart."
                )
                if not resp:
                    self.status_var.set(f"Update skipped ({latest}) — {now_ts()}")
                    return

                self._download_and_apply_update_async(info)

            self.after(0, ui)

        threading.Thread(target=worker, daemon=True).start()

    def _download_and_apply_update_async(self, info: dict):
        """
        Robust updater:
        - download asset (ZIP preferred; EXE supported)
        - validate (size if provided; signature)
        - extract if ZIP; locate Copy2.exe
        - stage into temp workdir
        - run updater .bat (backup + swap + rollback if new exe fails to start)
        """
        asset_url = str(info.get("asset_url", "")).strip()
        latest = str(info.get("version", "")).strip()
        html_url = str(info.get("html_url", "")).strip()
        expected_size = int(info.get("asset_size", 0) or 0)

        if not asset_url or not latest:
            messagebox.showwarning(APP_NAME, "Missing update download information.")
            return

        if not self._is_frozen():
            messagebox.showinfo(
                APP_NAME,
                "Auto-update is supported for the packaged portable .exe.\n\n"
                "You are running a Python script environment; opening the release page instead."
            )
            if html_url:
                webbrowser.open(html_url)
            return

        target_exe = Path(sys.executable)
        target_dir = target_exe.parent

        if not target_exe.exists():
            messagebox.showerror(APP_NAME, "Could not locate the current executable for updating.")
            return

        def worker():
            try:
                self._log_update_install(f"Update start -> latest={latest} url={asset_url}")
                self._log_update_install(f"Target exe: {target_exe}")

                workdir = Path(tempfile.mkdtemp(prefix="copy2_update_"))
                self._log_update_install(f"Workdir: {workdir}")

                # Download bytes
                req = urllib.request.Request(
                    asset_url,
                    headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"},
                    method="GET",
                )
                with urllib.request.urlopen(req, timeout=40) as r:
                    data = r.read()

                dl_path = workdir / "download.bin"
                dl_path.write_bytes(data)
                got_size = dl_path.stat().st_size
                self._log_update_install(f"Downloaded bytes: {got_size} expected={expected_size}")

                if expected_size > 0 and got_size != expected_size:
                    raise RuntimeError(f"Download size mismatch (got {got_size}, expected {expected_size}).")

                # Decide if ZIP or EXE
                head4 = data[:4]
                is_zip = (head4 == b"PK\x03\x04")
                is_exe = (data[:2] == b"MZ")

                staged_copy2 = workdir / "Copy2_new.exe"
                staged_uninst = workdir / "Copy2_Uninstall_new.exe"
                # If no uninstaller shipped, we'll leave staged_uninst missing/empty and batch will skip

                if is_zip:
                    zip_path = workdir / "update.zip"
                    shutil.move(str(dl_path), str(zip_path))

                    # Extract
                    extract_dir = workdir / "unzipped"
                    extract_dir.mkdir(parents=True, exist_ok=True)
                    with zipfile.ZipFile(zip_path, "r") as z:
                        z.extractall(extract_dir)

                    # Find Copy2.exe
                    new_copy2 = None
                    for p in extract_dir.rglob("Copy2.exe"):
                        new_copy2 = p
                        break

                    if new_copy2 is None:
                        # fallback: largest exe
                        exes = list(extract_dir.rglob("*.exe"))
                        if exes:
                            new_copy2 = sorted(exes, key=lambda x: x.stat().st_size, reverse=True)[0]

                    if new_copy2 is None or not new_copy2.exists():
                        raise RuntimeError("Could not locate Copy2.exe inside the update ZIP.")

                    if new_copy2.stat().st_size < MIN_EXE_BYTES:
                        raise RuntimeError("Extracted Copy2.exe is unexpectedly small; update package likely corrupt.")

                    shutil.copy2(new_copy2, staged_copy2)
                    self._log_update_install(f"Staged exe from zip: {new_copy2} size={new_copy2.stat().st_size}")

                    # Optional uninstaller (only if present)
                    found_uninst = None
                    for p in extract_dir.rglob("Copy2_Uninstall.exe"):
                        found_uninst = p
                        break
                    if found_uninst and found_uninst.exists():
                        shutil.copy2(found_uninst, staged_uninst)
                        self._log_update_install(f"Staged uninstaller: {found_uninst} size={found_uninst.stat().st_size}")
                    else:
                        # ensure not present
                        try:
                            if staged_uninst.exists():
                                staged_uninst.unlink()
                        except Exception:
                            pass

                elif is_exe:
                    # Direct exe asset
                    if got_size < MIN_EXE_BYTES:
                        raise RuntimeError("Downloaded EXE is unexpectedly small; update likely corrupt.")
                    shutil.move(str(dl_path), str(staged_copy2))
                    self._log_update_install(f"Staged exe direct: {staged_copy2} size={staged_copy2.stat().st_size}")
                    # No uninstaller
                    try:
                        if staged_uninst.exists():
                            staged_uninst.unlink()
                    except Exception:
                        pass

                else:
                    # Probably HTML/error page saved
                    preview = dl_path.read_text(encoding="utf-8", errors="ignore")[:240]
                    self._log_update_install(f"Signature mismatch. First chars:\n{preview}")
                    raise RuntimeError("Downloaded file is not a ZIP or EXE. (Possible GitHub HTML/error response.)")

                # Build updater batch (backup + swap + rollback)
                pid = os.getpid()
                exe_name = target_exe.name
                bak_exe = target_dir / (exe_name + ".bak")
                bat = workdir / "copy2_updater.bat"

                bat_contents = f"""@echo off
setlocal enabledelayedexpansion

set PID={pid}
set TARGET_EXE="{str(target_exe)}"
set TARGET_DIR="{str(target_dir)}"
set EXE_NAME={exe_name}
set BAK_EXE="{str(bak_exe)}"
set NEW_COPY2="{str(staged_copy2)}"
set NEW_UNINST="{str(staged_uninst)}"

:wait
for /f "tokens=2 delims=," %%A in ('tasklist /FI "PID eq %PID%" /FO CSV /NH 2^>NUL') do (
  if "%%~A"=="%PID%" (
    timeout /t 1 /nobreak >NUL
    goto wait
  )
)

REM Backup current exe
if exist %BAK_EXE% del /f /q %BAK_EXE% >NUL 2>&1
copy /y %TARGET_EXE% %BAK_EXE% >NUL 2>&1

REM Swap in new exe
move /y %NEW_COPY2% %TARGET_EXE% >NUL 2>&1

REM Optional uninstaller swap if provided
if exist %NEW_UNINST% (
  for %%I in (%NEW_UNINST%) do set SIZE=%%~zI
  if NOT "!SIZE!"=="0" (
    move /y %NEW_UNINST% %TARGET_DIR%\\Copy2_Uninstall.exe >NUL 2>&1
  )
)

REM Start updated app
start "" /d %TARGET_DIR% %TARGET_EXE%
timeout /t 2 /nobreak >NUL

REM If it didn't start, rollback
tasklist /FI "IMAGENAME eq %EXE_NAME%" /NH | find /I "%EXE_NAME%" >NUL
if errorlevel 1 (
  copy /y %BAK_EXE% %TARGET_EXE% >NUL 2>&1
  start "" /d %TARGET_DIR% %TARGET_EXE%
)

del "%~f0"
endlocal
"""
                bat.write_text(bat_contents, encoding="utf-8", errors="ignore")
                self._log_update_install(f"Wrote updater bat: {bat}")

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

                # Close this instance
                self.after(0, lambda: self._on_close_for_update())

            except Exception as e:
                self._log_update_install(f"ERROR: {repr(e)}")
                def ui_err():
                    messagebox.showerror(
                        APP_NAME,
                        f"Update failed:\n{e}\n\nLog:\n{self.data_dir / 'update_install.log'}"
                    )
                    self.status_var.set(f"Update failed — {now_ts()}")
                self.after(0, ui_err)

        threading.Thread(target=worker, daemon=True).start()

    def _on_close_for_update(self):
        # Like normal close, but no “unsaved edits” blocking
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

            # Start clipboard polling
            self.after(250, self._poll_clipboard)

            # Update check on startup
            if CHECK_UPDATES_ON_STARTUP:
                self.after(STARTUP_UPDATE_DELAY_MS, lambda: self._check_updates_async(prompt_if_new=True))

            self.protocol("WM_DELETE_WINDOW", self._on_close)

        def _apply_window_defaults(self):
            self.minsize(1100, 720)
            try:
                self.tk.call("tk", "scaling", 1.2)
            except Exception:
                pass

        def _tooltip(self, widget, text: str):
            if not text:
                return
            try:
                if TBToolTip:
                    TBToolTip(widget, text=text)
                else:
                    SimpleTooltip(widget, text)
            except Exception:
                try:
                    SimpleTooltip(widget, text)
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

            btn_find = tb.Button(search_box, text="Find", command=self._search, bootstyle=PRIMARY)
            btn_find.pack(side=LEFT, padx=(10, 0))
            self._tooltip(btn_find, "Find (Enter)")

            btn_prev = tb.Button(search_box, text="Prev", command=lambda: self._jump_match(-1), bootstyle=SECONDARY)
            btn_prev.pack(side=LEFT, padx=(8, 0))
            btn_next = tb.Button(search_box, text="Next", command=lambda: self._jump_match(1), bootstyle=SECONDARY)
            btn_next.pack(side=LEFT, padx=(8, 0))

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
            self._tooltip(pause_btn, "Pause/Resume capturing")

            upd_btn = tb.Button(action_box, text="Check Updates", command=lambda: self._check_updates_async(prompt_if_new=True), bootstyle=OUTLINE)
            upd_btn.pack(side=LEFT, padx=(0, 8))
            self._tooltip(upd_btn, "Check for updates (GitHub Releases)")

            exp_btn = tb.Button(action_box, text="Export", command=self._export, bootstyle=OUTLINE)
            exp_btn.pack(side=LEFT, padx=(0, 8))
            self._tooltip(exp_btn, "Export (Ctrl+E)")

            imp_btn = tb.Button(action_box, text="Import", command=self._import, bootstyle=OUTLINE)
            imp_btn.pack(side=LEFT, padx=(0, 8))
            self._tooltip(imp_btn, "Import (Ctrl+I)")

            set_btn = tb.Button(action_box, text="Settings", command=self._open_settings, bootstyle=SECONDARY)
            set_btn.pack(side=LEFT)
            self._tooltip(set_btn, "Settings / Controls")

            tb.Separator(root).pack(fill=X, pady=(0, 10))

            # Body split
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
            tb.Radiobutton(filters, text="All", value="all", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=0, padx=(0, 8))
            tb.Radiobutton(filters, text="Favorites", value="fav", variable=self.filter_var, command=self._refresh_list, bootstyle="toolbutton").grid(row=0, column=1, padx=(0, 8))

            clean_btn = tb.Button(filters, text="Clean", command=self._clean_history, bootstyle=DANGER)
            clean_btn.grid(row=0, column=4, sticky="e")
            self._tooltip(clean_btn, "Clean non-favorites (Ctrl+L)")

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

            b_copy = tb.Button(left_actions, text="Copy", command=self._copy_selected, bootstyle=SUCCESS)
            b_copy.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            self._tooltip(b_copy, "Copy (Ctrl+C)")

            b_del = tb.Button(left_actions, text="Delete", command=self._delete_selected, bootstyle=WARNING)
            b_del.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            self._tooltip(b_del, "Delete (Del)")

            b_fav = tb.Button(left_actions, text="Fav / Unfav", command=self._toggle_favorite_selected, bootstyle=INFO)
            b_fav.grid(row=0, column=2, sticky="ew", padx=(0, 8))

            b_comb = tb.Button(left_actions, text="Combine", command=self._combine_selected, bootstyle=PRIMARY)
            b_comb.grid(row=0, column=3, sticky="ew")
            self._tooltip(b_comb, "Combine selected (Ctrl+Click multi-select)")

            # Right pane
            right = tb.Labelframe(body, text="Preview (Editable)", padding=10)
            right.grid(row=0, column=1, sticky="nsew")
            right.rowconfigure(1, weight=1)
            right.columnconfigure(0, weight=1)

            preview_actions = tb.Frame(right)
            preview_actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            preview_actions.columnconfigure(0, weight=1)

            self.reverse_var = tk.BooleanVar(value=self.settings.reverse_lines_copy)
            rev_toggle = tb.Checkbutton(
                preview_actions,
                text="Reverse-lines copy",
                variable=self.reverse_var,
                command=self._toggle_reverse_lines,
                bootstyle="round-toggle",
            )
            rev_toggle.pack(side=LEFT)
            self._tooltip(rev_toggle, "Reverse lines (preview shows what will be copied)")

            self.preview_dirty_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.preview_dirty_var).pack(side=LEFT, padx=(10, 0))

            self.revert_btn = tb.Button(preview_actions, text="Revert", command=self._revert_preview_edits, bootstyle=WARNING)
            self.revert_btn.pack(side=RIGHT, padx=(0, 8))
            self._tooltip(self.revert_btn, "Revert (Esc)")

            self.save_btn = tb.Button(preview_actions, text="Save Edit", command=self._save_preview_edits, bootstyle=SUCCESS)
            self.save_btn.pack(side=RIGHT, padx=(0, 8))
            self._tooltip(self.save_btn, "Save (Ctrl+S)")

            copy_prev = tb.Button(preview_actions, text="Copy Preview", command=self._copy_selected, bootstyle=SUCCESS)
            copy_prev.pack(side=RIGHT)
            self._tooltip(copy_prev, "Copy Preview (Ctrl+C)")

            self.preview = tk.Text(right, wrap="word", undo=True)
            self.preview.grid(row=1, column=0, sticky="nsew")
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            sb2 = tb.Scrollbar(right, orient="vertical", command=self.preview.yview)
            sb2.grid(row=1, column=1, sticky="ns")
            self.preview.configure(yscrollcommand=sb2.set)

            bottom = tb.Frame(root)
            bottom.pack(fill=X, pady=(10, 0))
            self.status_var = tk.StringVar(value=f"Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}")
            tb.Label(bottom, textvariable=self.status_var).pack(side=LEFT)

            self._update_preview_dirty_ui()

        def _bind_shortcuts(self):
            self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
            self.bind("<Return>", lambda e: (self._search(), "break"))
            self.bind("<Control-c>", lambda e: (self._copy_selected(), "break"))
            self.bind("<Delete>", lambda e: (self._delete_selected(), "break"))
            self.bind("<Control-e>", lambda e: (self._export(), "break"))
            self.bind("<Control-i>", lambda e: (self._import(), "break"))
            self.bind("<Control-l>", lambda e: (self._clean_history(), "break"))
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

            self._persist()
            self.destroy()


else:
    # Minimal ttk fallback
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
            if CHECK_UPDATES_ON_STARTUP:
                self.after(STARTUP_UPDATE_DELAY_MS, lambda: self._check_updates_async(prompt_if_new=True))
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
            self.status_var = tk.StringVar(value="")

            self.preview_dirty_var = tk.StringVar(value="")
            self.save_btn = None
            self.revert_btn = None

        def _bind_shortcuts(self):
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))

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
