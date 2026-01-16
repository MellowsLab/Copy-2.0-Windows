"""
Copy 2.0 (Windows Portable) - single-file Tkinter clipboard history tool.

Notes:
- Designed for packaging into a single .exe via PyInstaller (--onefile --noconsole).
- Clipboard capture uses polling (pyperclip) for broad compatibility.
- Global hotkeys are intentionally not implemented by default because Windows-global
  hooks can require elevated privileges and sometimes trigger AV heuristics. The app
  includes in-app shortcuts (Ctrl+F, Ctrl+C, Del, Ctrl+E, Ctrl+I, etc.).
"""

from __future__ import annotations

import json
import re
from collections import deque
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pyperclip
from platformdirs import user_data_dir


APP_NAME = "Copy 2.0"
APP_ID = "copy2"
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
    wrap_mode: bool = False  # reverse-lines copy toggle

    @staticmethod
    def from_dict(d: dict) -> "Settings":
        s = Settings()
        s.max_history = int(d.get("max_history", DEFAULT_MAX_HISTORY))
        s.poll_ms = int(d.get("poll_ms", DEFAULT_POLL_MS))
        s.session_only = bool(d.get("session_only", False))
        s.wrap_mode = bool(d.get("wrap_mode", False))
        s.max_history = max(5, min(500, s.max_history))
        s.poll_ms = max(100, min(5000, s.poll_ms))
        return s


class Copy2App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.minsize(980, 620)

        self.data_dir = Path(user_data_dir(APP_ID, "MellowsLab"))
        self.settings_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.favs_path = self.data_dir / "favorites.json"

        self.settings = Settings.from_dict(safe_json_load(self.settings_path, {}))
        self.history = deque(safe_json_load(self.history_path, []), maxlen=self.settings.max_history)
        self.favorites = safe_json_load(self.favs_path, [])

        self.paused = False
        self.last_clip = ""
        self.search_matches: list[int] = []
        self.search_index = 0
        self._poll_job = None

        self._build_ui()
        self._bind_shortcuts()
        self._refresh_list()

        self.after(250, self._poll_clipboard)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- UI ----------------
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=10)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=2)
        root.rowconfigure(1, weight=1)

        top = ttk.Frame(root)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        top.columnconfigure(2, weight=1)

        self.pause_var = tk.BooleanVar(value=self.paused)
        ttk.Checkbutton(top, text="Pause capture", variable=self.pause_var, command=self._toggle_pause).grid(
            row=0, column=0, sticky="w"
        )

        ttk.Label(top, text="Search:").grid(row=0, column=1, sticky="e", padx=(10, 4))
        self.search_var = tk.StringVar(value="")
        self.search_entry = ttk.Entry(top, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=2, sticky="ew")

        ttk.Button(top, text="Find", command=self._search).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(top, text="Prev", command=lambda: self._jump_match(-1)).grid(row=0, column=4, padx=(6, 0))
        ttk.Button(top, text="Next", command=lambda: self._jump_match(1)).grid(row=0, column=5, padx=(6, 0))

        left = ttk.Frame(root)
        left.grid(row=1, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        ttk.Label(left, text="History").grid(row=0, column=0, sticky="w")

        self.listbox = tk.Listbox(left, activestyle="dotbox", exportselection=False)
        self.listbox.grid(row=1, column=0, sticky="nsew")
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._on_select())

        sb = ttk.Scrollbar(left, orient="vertical", command=self.listbox.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=sb.set)

        left_btns = ttk.Frame(left)
        left_btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for i in range(6):
            left_btns.columnconfigure(i, weight=1)

        ttk.Button(left_btns, text="Copy", command=self._copy_selected).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(left_btns, text="Delete", command=self._delete_selected).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(left_btns, text="Clear", command=self._clear_history).grid(row=0, column=2, sticky="ew", padx=(0, 6))
        ttk.Button(left_btns, text="Fav/Unfav", command=self._toggle_favorite_selected).grid(row=0, column=3, sticky="ew", padx=(0, 6))
        ttk.Button(left_btns, text="Combine", command=self._combine_selected).grid(row=0, column=4, sticky="ew", padx=(0, 6))
        ttk.Button(left_btns, text="Settings", command=self._open_settings).grid(row=0, column=5, sticky="ew")

        right = ttk.Frame(root)
        right.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        ttk.Label(right, text="Preview").grid(row=0, column=0, sticky="w")

        self.preview = tk.Text(right, wrap="word")
        self.preview.grid(row=1, column=0, sticky="nsew")
        self.preview.configure(state="disabled")

        sb2 = ttk.Scrollbar(right, orient="vertical", command=self.preview.yview)
        sb2.grid(row=1, column=1, sticky="ns")
        self.preview.configure(yscrollcommand=sb2.set)

        bottom = ttk.Frame(root)
        bottom.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        bottom.columnconfigure(2, weight=1)

        ttk.Button(bottom, text="Export…", command=self._export).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(bottom, text="Import…", command=self._import).grid(row=0, column=1, padx=(0, 6))

        self.wrap_var = tk.BooleanVar(value=self.settings.wrap_mode)
        ttk.Checkbutton(bottom, text="Reverse-lines copy", variable=self.wrap_var, command=self._toggle_wrap_mode).grid(
            row=0, column=2, sticky="w"
        )

        self.status_var = tk.StringVar(value=f"Loaded {len(self.history)} items. Data: {self.data_dir}")
        ttk.Label(bottom, textvariable=self.status_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _bind_shortcuts(self):
        self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
        self.bind("<Return>", lambda e: (self._search(), "break"))
        self.bind("<Control-c>", lambda e: (self._copy_selected(), "break"))
        self.bind("<Delete>", lambda e: (self._delete_selected(), "break"))
        self.bind("<Control-e>", lambda e: (self._export(), "break"))
        self.bind("<Control-i>", lambda e: (self._import(), "break"))
        self.bind("<Control-l>", lambda e: (self._clear_history(), "break"))
        self.bind("<Escape>", lambda e: (self.search_var.set(""), "break"))

    # ---------------- Clipboard polling ----------------
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

    # ---------------- Persistence ----------------
    def _persist(self):
        if self.settings.session_only:
            return
        safe_json_save(self.settings_path, asdict(self.settings))
        safe_json_save(self.history_path, list(self.history))
        safe_json_save(self.favs_path, list(self.favorites))

    # ---------------- List/preview helpers ----------------
    def _refresh_list(self, select_last: bool = False):
        self.listbox.delete(0, tk.END)
        for item in self.history:
            self.listbox.insert(tk.END, self._format_list_item(item))
        if select_last and len(self.history) > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(tk.END)
            self.listbox.see(tk.END)
            self._on_select()

        self.status_var.set(f"Loaded {len(self.history)} items. Favorites: {len(self.favorites)}. Data: {self.data_dir}")

    def _format_list_item(self, item: str) -> str:
        one = re.sub(r"\s+", " ", item).strip()
        if len(one) > 90:
            one = one[:87] + "..."
        prefix = "★ " if item in self.favorites else "  "
        return prefix + one

    def _get_selected_index(self):
        sel = self.listbox.curselection()
        return int(sel[0]) if sel else None

    def _get_selected_text(self):
        i = self._get_selected_index()
        if i is None:
            return None
        try:
            return list(self.history)[i]
        except Exception:
            return None

    def _on_select(self):
        t = self._get_selected_text() or ""
        self.preview.configure(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", t)
        self.preview.configure(state="disabled")

    # ---------------- Actions ----------------
    def _toggle_pause(self):
        self.paused = bool(self.pause_var.get())
        self.status_var.set("Capture paused." if self.paused else "Capture running.")

    def _toggle_wrap_mode(self):
        self.settings.wrap_mode = bool(self.wrap_var.get())
        self._persist()

    def _copy_selected(self):
        t = self._get_selected_text()
        if not t:
            return
        out = t
        if self.settings.wrap_mode:
            lines = out.splitlines()
            lines.reverse()
            out = "\n".join(lines)
        try:
            pyperclip.copy(out)
            self.status_var.set(f"Copied to clipboard at {now_ts()}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to copy to clipboard:\n{e}")

    def _delete_selected(self):
        i = self._get_selected_index()
        if i is None:
            return
        items = list(self.history)
        removed = items.pop(i)
        self.history = deque(items, maxlen=self.settings.max_history)
        if removed in self.favorites:
            self.favorites = [x for x in self.favorites if x != removed]
        self._refresh_list(select_last=True)
        self._persist()

    def _clear_history(self):
        if not messagebox.askyesno(APP_NAME, "Clear all history items?"):
            return
        self.history = deque([], maxlen=self.settings.max_history)
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
        sels = list(self.listbox.curselection())
        if not sels:
            return
        items = list(self.history)
        combined = "\n".join(items[i] for i in sels if 0 <= i < len(items))
        if combined:
            try:
                pyperclip.copy(combined)
                self.status_var.set(f"Combined {len(sels)} items and copied.")
            except Exception as e:
                messagebox.showerror(APP_NAME, f"Failed to copy combined text:\n{e}")

    # ---------------- Search ----------------
    def _search(self):
        q = self.search_var.get().strip()
        self.search_matches = []
        self.search_index = 0
        if not q:
            self.status_var.set("Search cleared.")
            return
        items = list(self.history)
        for i, item in enumerate(items):
            if q.lower() in item.lower():
                self.search_matches.append(i)
        if not self.search_matches:
            self.status_var.set(f"No matches for: {q}")
            return
        self.status_var.set(f"Found {len(self.search_matches)} matches for: {q}")
        self._jump_to_index(self.search_matches[0])

    def _jump_to_index(self, idx: int):
        if idx < 0 or idx >= len(self.history):
            return
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.see(idx)
        self._on_select()

    def _jump_match(self, direction: int):
        if not self.search_matches:
            self._search()
            return
        self.search_index = (self.search_index + direction) % len(self.search_matches)
        self._jump_to_index(self.search_matches[self.search_index])

    # ---------------- Import/Export ----------------
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
            self.status_var.set(f"Exported to {path}")
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
                seen = set()
                out = []
                for item in reversed(merged):
                    if item not in seen:
                        out.append(item)
                        seen.add(item)
                out.reverse()
                self.history = deque(out[-self.settings.max_history :], maxlen=self.settings.max_history)
            if isinstance(favorites, list):
                self.favorites = [x for x in favorites if isinstance(x, str)]
            self._refresh_list(select_last=True)
            self._persist()
            self.status_var.set(f"Imported from {path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")

    # ---------------- Settings dialog ----------------
    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        max_var = tk.StringVar(value=str(self.settings.max_history))
        poll_var = tk.StringVar(value=str(self.settings.poll_ms))
        sess_var = tk.BooleanVar(value=self.settings.session_only)

        ttk.Label(frm, text="Max history (5–500):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=max_var, width=12).grid(row=0, column=1, sticky="w", padx=(8, 0))

        ttk.Label(frm, text="Poll interval ms (100–5000):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frm, textvariable=poll_var, width=12).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        ttk.Checkbutton(frm, text="Session-only (do not save history)", variable=sess_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).grid(row=0, column=0, padx=(0, 8))

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
            items = list(self.history)[-mh:]
            self.history = deque(items, maxlen=mh)
            self._refresh_list(select_last=True)
            self._persist()
            dlg.destroy()

        ttk.Button(btns, text="Save", command=save).grid(row=0, column=1)

    # ---------------- Close ----------------
    def _on_close(self):
        try:
            if self._poll_job is not None:
                self.after_cancel(self._poll_job)
        except Exception:
            pass
        self._persist()
        self.destroy()


def main():
    Copy2App().mainloop()


if __name__ == "__main__":
    main()
