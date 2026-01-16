"""
Copy 2.0 (Windows) — Modern UI Edition (Improved)

Drop-in replacement for Copy2_Windows.py.

Changes vs previous:
- Reverse-lines copy visibly reverses Preview (what you see is what you paste).
- Preview is editable + Save Edit (Ctrl+S) + Revert (Esc).
- Combine supports Ctrl+Click multi-select in a specific order:
  creates a NEW history item, leaves originals intact, selects & previews the new item.
- Find selects matching history item and highlights the query inside Preview.
- Slightly more spacing for Pause toggle.

Dependencies:
  pyperclip
  platformdirs
Recommended for modern UI:
  ttkbootstrap

Build:
  python -m PyInstaller --noconsole --onefile --name "Copy2" Copy2_Windows.py
"""

from __future__ import annotations

import json
import re
from collections import deque
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

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


APP_NAME = "Copy 2.0"
APP_ID = "copy2"
VENDOR = "MellowsLab"

DEFAULT_MAX_HISTORY = 50
DEFAULT_POLL_MS = 400


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
        s.max_history = max(5, min(500, s.max_history))
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
        self.history = deque(safe_json_load(self.history_path, []), maxlen=self.settings.max_history)
        self.favorites = safe_json_load(self.favs_path, [])

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

    def _persist(self):
        if self.settings.session_only:
            return
        safe_json_save(self.settings_path, asdict(self.settings))
        safe_json_save(self.history_path, list(self.history))
        safe_json_save(self.favs_path, list(self.favorites))

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

        self.history = deque(items, maxlen=self.settings.max_history)
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
        # Editable preview: keep state normal
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", text)
        self._preview_dirty = not mark_clean
        self._update_preview_dirty_ui()

        # Re-highlight search query in preview, if any
        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _on_select(self):
        # If the user has unsaved edits and changes selection, ask what to do
        if self._preview_dirty and self._selected_item_text is not None:
            resp = messagebox.askyesnocancel(
                APP_NAME,
                "You have unsaved edits in Preview.\n\n"
                "Yes = Save edits\nNo = Discard edits\nCancel = Stay on current item"
            )
            if resp is None:
                # Cancel: revert list selection back to previously selected item (best-effort)
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

        # Store original selected item text (non-reversed)
        self._selected_item_text = t
        display = self._get_preview_display_text_for_item(t)
        self._set_preview_text(display, mark_clean=True)

        # Ensure highlight if searching
        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _reselect_current_item(self):
        # Reselect the item matching _selected_item_text (if possible)
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

        # Remove deselected indices from order
        if removed:
            self._sel_order = [i for i in self._sel_order if i not in removed]

        # Add new selections in a stable way
        # Typically ctrl-click adds one index -> preserve that click order.
        # If a range adds multiple at once, append in ascending order.
        if added:
            for i in sorted(list(added)):
                if i not in self._sel_order:
                    self._sel_order.append(i)

        self._prev_sel_set = current

        # Keep preview synced with primary selection (first selected)
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

        # Re-render preview for current selection (what you see is what you paste)
        if self._selected_item_text is not None:
            display = self._get_preview_display_text_for_item(self._selected_item_text)
            # If user had edits, do not overwrite blindly—ask.
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
        # Only mark dirty if we have a selected item (or if user is editing a combined/new view)
        if self._selected_item_text is None and self.preview.get("1.0", tk.END).strip():
            self._preview_dirty = True
        elif self._selected_item_text is not None:
            # Compare against displayed baseline
            baseline = self._get_preview_display_text_for_item(self._selected_item_text)
            current = self.preview.get("1.0", tk.END).rstrip("\n")
            self._preview_dirty = (current != baseline)
        self._update_preview_dirty_ui()

        # Maintain search highlight while editing
        if self.search_query:
            self._highlight_query_in_preview(self.search_query)

    def _save_preview_edits(self, _event=None):
        text_now = self.preview.get("1.0", tk.END).rstrip("\n")

        # If no selected item, save as a new history entry
        if self._selected_item_text is None:
            if text_now.strip():
                self._add_history_item(text_now)
                self.status_var.set(f"Saved new item from Preview — {now_ts()}")
                self._preview_dirty = False
                self._update_preview_dirty_ui()
            return

        # Replace the selected item in history with the edited text.
        # Note: If reverse-lines is ON, we are saving exactly what is displayed (what you intended to paste).
        old = self._selected_item_text
        new = text_now

        if not new.strip():
            messagebox.showwarning(APP_NAME, "Cannot save an empty item.")
            return

        items = [x for x in self.history if x != old]
        items.append(new)
        self.history = deque(items, maxlen=self.settings.max_history)

        # Update favorites mapping if needed
        if old in self.favorites:
            self.favorites = [new if x == old else x for x in self.favorites]

        self._selected_item_text = new  # new selection baseline
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
        # Copy exactly what is shown in Preview (WYSIWYG)
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

        # Clear preview if we deleted the selected baseline
        if self._selected_item_text == t:
            self._selected_item_text = None
            self._set_preview_text("", mark_clean=True)

        self._refresh_list(select_last=True)
        self._persist()

    def _clear_history(self):
        if not messagebox.askyesno(APP_NAME, "Clear all clipboard history items?"):
            return
        self.history = deque([], maxlen=self.settings.max_history)
        self._selected_item_text = None
        self._set_preview_text("", mark_clean=True)
        self._refresh_list()
        self._persist()

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

        # Use click-order where possible
        ordered = [i for i in self._sel_order if i in sel]
        # If for some reason order list is empty, fall back to current selection order
        if not ordered:
            ordered = sel

        parts = []
        for i in ordered:
            if 0 <= i < len(self.view_items):
                parts.append(self.view_items[i])

        if not parts:
            return

        combined = "\n".join(parts)

        # Create a NEW history item, keep originals intact
        self._add_history_item(combined)

        # Select the new item (it will be the latest)
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

        # Find first occurrence (case-insensitive)
        pattern = re.compile(re.escape(query), re.IGNORECASE)
        m = pattern.search(text)
        if not m:
            return

        start_index = f"1.0+{m.start()}c"
        end_index = f"1.0+{m.end()}c"

        self.preview.tag_add("match", start_index, end_index)
        # Configure tag style
        try:
            self.preview.tag_config("match", background="#2b78ff", foreground="white")
        except Exception:
            pass

        # Scroll to it
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

        # Update selection baseline + preview
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

            if isinstance(history, list):
                merged = list(self.history)
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

                self.history = deque(out[-self.settings.max_history:], maxlen=self.settings.max_history)

            if isinstance(favorites, list):
                self.favorites = [x for x in favorites if isinstance(x, str)]

            self._refresh_list(select_last=True)
            self._persist()
            self.status_var.set(f"Imported — {path}")

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")

    # -----------------------------
    # Settings dialog (unchanged)
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
        else:
            from tkinter import ttk as _ttk  # type: ignore
            Frame = _ttk.Frame
            Label = _ttk.Label
            Entry = _ttk.Entry
            Button = _ttk.Button
            Checkbutton = _ttk.Checkbutton
            Combobox = _ttk.Combobox

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

        btns = Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))

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

            items = list(self.history)[-mh:]
            self.history = deque(items, maxlen=mh)

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
        self.menu.add_command(label="Favorite / Unfavorite", command=self._toggle_favorite_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Save Preview Edit", command=self._save_preview_edits)
        self.menu.add_command(label="Revert Preview Edit", command=self._revert_preview_edits)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self._delete_selected)
        self.menu.add_command(label="Combine Selected", command=self._combine_selected)

        def popup(event):
            try:
                idx = self.listbox.nearest(event.y)
                if idx >= 0:
                    # Keep multi-select intact if Ctrl is held
                    if not (event.state & 0x0004):  # Control key mask
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

            tb.Button(search_box, text="Find", command=self._search, bootstyle=PRIMARY).pack(side=LEFT, padx=(10, 0))
            tb.Button(search_box, text="Prev", command=lambda: self._jump_match(-1), bootstyle=SECONDARY).pack(side=LEFT, padx=(8, 0))
            tb.Button(search_box, text="Next", command=lambda: self._jump_match(1), bootstyle=SECONDARY).pack(side=LEFT, padx=(8, 0))

            # Spacer to ensure Pause isn't too close
            tb.Frame(top, width=18).pack(side=LEFT)

            # Right actions
            action_box = tb.Frame(top)
            action_box.pack(side=RIGHT)

            self.pause_var = tk.BooleanVar(value=self.paused)
            tb.Checkbutton(
                action_box,
                text="Pause",
                variable=self.pause_var,
                command=self._toggle_pause,
                bootstyle="round-toggle",
            ).pack(side=LEFT, padx=(10, 18))

            tb.Button(action_box, text="Export", command=self._export, bootstyle=OUTLINE).pack(side=LEFT, padx=(0, 8))
            tb.Button(action_box, text="Import", command=self._import, bootstyle=OUTLINE).pack(side=LEFT, padx=(0, 8))
            tb.Button(action_box, text="Settings", command=self._open_settings, bootstyle=SECONDARY).pack(side=LEFT)

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
            tb.Button(filters, text="Clear", command=self._clear_history, bootstyle=DANGER).grid(row=0, column=4, sticky="e")

            # IMPORTANT: Extended select for Ctrl+Click multi-select
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

            tb.Button(left_actions, text="Copy", command=self._copy_selected, bootstyle=SUCCESS).grid(row=0, column=0, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Delete", command=self._delete_selected, bootstyle=WARNING).grid(row=0, column=1, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Fav / Unfav", command=self._toggle_favorite_selected, bootstyle=INFO).grid(row=0, column=2, sticky="ew", padx=(0, 8))
            tb.Button(left_actions, text="Combine", command=self._combine_selected, bootstyle=PRIMARY).grid(row=0, column=3, sticky="ew")

            # Right pane (Preview)
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

            # Dirty indicator
            self.preview_dirty_var = tk.StringVar(value="")
            tb.Label(preview_actions, textvariable=self.preview_dirty_var).pack(side=LEFT, padx=(10, 0))

            # Buttons on right
            self.revert_btn = tb.Button(preview_actions, text="Revert", command=self._revert_preview_edits, bootstyle=WARNING)
            self.revert_btn.pack(side=RIGHT, padx=(0, 8))
            self.save_btn = tb.Button(preview_actions, text="Save Edit", command=self._save_preview_edits, bootstyle=SUCCESS)
            self.save_btn.pack(side=RIGHT, padx=(0, 8))

            tb.Button(preview_actions, text="Copy Preview", command=self._copy_selected, bootstyle=SUCCESS).pack(side=RIGHT)

            # Editable preview text
            self.preview = tk.Text(right, wrap="word", undo=True)
            self.preview.grid(row=1, column=0, sticky="nsew")
            self.preview.bind("<KeyRelease>", self._mark_preview_dirty)

            sb2 = tb.Scrollbar(right, orient="vertical", command=self.preview.yview)
            sb2.grid(row=1, column=1, sticky="ns")
            self.preview.configure(yscrollcommand=sb2.set)

            # Bottom status bar
            bottom = tb.Frame(root)
            bottom.pack(fill=X, pady=(10, 0))
            self.status_var = tk.StringVar(value=f"Items: {len(self.history)}   Favorites: {len(self.favorites)}   Data: {self.data_dir}")
            tb.Label(bottom, textvariable=self.status_var).pack(side=LEFT)

            # Initialize preview buttons state
            self._update_preview_dirty_ui()

        def _bind_shortcuts(self):
            self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
            self.bind("<Return>", lambda e: (self._search(), "break"))
            self.bind("<Control-c>", lambda e: (self._copy_selected(), "break"))
            self.bind("<Delete>", lambda e: (self._delete_selected(), "break"))
            self.bind("<Control-e>", lambda e: (self._export(), "break"))
            self.bind("<Control-i>", lambda e: (self._import(), "break"))
            self.bind("<Control-l>", lambda e: (self._clear_history(), "break"))
            self.bind("<Control-s>", lambda e: (self._save_preview_edits(), "break"))
            self.bind("<Escape>", lambda e: (self._revert_preview_edits(), "break"))

        def _on_close(self):
            try:
                if self._poll_job is not None:
                    self.after_cancel(self._poll_job)
            except Exception:
                pass

            # If preview has unsaved changes on exit, ask
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
    # ttk fallback (kept minimal; modern UI requires ttkbootstrap)
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
