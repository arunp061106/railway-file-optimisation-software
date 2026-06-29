"""
app.py
------
Railway File Organiser — Main Desktop Application
A beautiful, officer-friendly GUI app. One double-click to run.
No technical knowledge required.
"""

import os
import sys
import json
import threading
import time
import queue
import shutil
import logging
import ctypes
import ctypes.wintypes
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── Windows API helpers for forcing popup to front ─────────────────────────────
def _flash_window(hwnd):
    """Flash the taskbar button to alert the user (Windows only)."""
    try:
        FLASHW_ALL       = 0x00000003
        FLASHW_TIMERNOFG = 0x0000000C
        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize",    ctypes.wintypes.UINT),
                ("hwnd",      ctypes.wintypes.HWND),
                ("dwFlags",   ctypes.wintypes.DWORD),
                ("uCount",    ctypes.wintypes.UINT),
                ("dwTimeout", ctypes.wintypes.DWORD),
            ]
        fw = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd,
                        FLASHW_ALL | FLASHW_TIMERNOFG, 8, 0)
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(fw))
    except Exception:
        pass

def _force_window_front(hwnd):
    """Force a window to the foreground using Windows API."""
    try:
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    except Exception:
        pass

# ── Optional tray ──────────────────────────────────────────────────────────────
try:
    import pystray
    from PIL import Image, ImageDraw, ImageFont
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ── Internal modules ───────────────────────────────────────────────────────────
BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from categoriser  import Categoriser
from file_watcher import FileWatcher
from file_mover   import FileMover
from excel_logger import ExcelLogger

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    filename=str(BASE_DIR / "organiser.log"),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# ── Config paths ───────────────────────────────────────────────────────────────
CONFIG_PATH     = BASE_DIR / "config.json"
CATEGORIES_PATH = BASE_DIR / "categories.json"

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN TOKENS
# ══════════════════════════════════════════════════════════════════════════════
C = {
    "bg":       "#0A0F1E",   # Deep space navy
    "surface":  "#131929",   # Card surface
    "surface2": "#1C2540",   # Elevated surface
    "border":   "#2A3557",   # Subtle border
    "accent":   "#2563EB",   # Railway blue
    "accent2":  "#1D4ED8",   # Darker blue
    "green":    "#16A34A",   # Success green
    "green2":   "#15803D",   # Darker green
    "yellow":   "#D97706",   # Warning amber
    "red":      "#DC2626",   # Error red
    "text":     "#F0F4FF",   # Primary text
    "text2":    "#8B9CBD",   # Secondary text
    "text3":    "#4A5880",   # Muted text
    "white":    "#FFFFFF",
}

FONT_TITLE  = ("Segoe UI", 22, "bold")
FONT_HEAD   = ("Segoe UI", 13, "bold")
FONT_BODY   = ("Segoe UI", 11)
FONT_SMALL  = ("Segoe UI", 9)
FONT_BADGE  = ("Segoe UI", 9, "bold")
FONT_MONO   = ("Consolas", 9)


# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
class Config:
    DEFAULTS = {
        "watch_folder": str(Path.home() / "Downloads"),
        "base_folder":  str(Path.home() / "Documents" / "Railway Files"),
        "excel_log_filename": "Railway_Files_Log.xlsx",
        "auto_confirm_timeout_seconds": 60,
        "open_excel_after_update": True,
        "stable_wait_seconds": 2,
    }

    def __init__(self):
        self._d = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, encoding="utf-8") as f:
                    self._d.update(json.load(f))
            except Exception:
                pass
        if not self._d["watch_folder"]:
            self._d["watch_folder"] = self.DEFAULTS["watch_folder"]
        if not self._d["base_folder"]:
            self._d["base_folder"] = self.DEFAULTS["base_folder"]

    def save(self):
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self._d, f, indent=4)
        except Exception as e:
            logging.error(f"Config save failed: {e}")

    def __getitem__(self, k):  return self._d.get(k, self.DEFAULTS.get(k))
    def __setitem__(self, k, v): self._d[k] = v

    def log_path(self):
        return str(Path(self._d["base_folder"]) / self._d["excel_log_filename"])


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class RailwayApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.cfg  = Config()
        self.cat  = Categoriser(CATEGORIES_PATH)
        self.watcher = None
        self.watching = False
        self._ui_queue = queue.Queue()   # thread-safe UI updates
        self._pending_file = None        # file waiting for user confirm
        self._row_paths = {}             # treeview row_id -> file_path

        self._setup_styles()
        self._build_window()
        self._build_ui()
        self._poll_ui_queue()
        self._refresh_status()

        # Auto-start watching after UI is ready (500ms delay)
        self.after(500, self._load_and_verify_register)
        self.after(800, self._auto_start_if_configured)

    def _setup_styles(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
            
        # Configure Notebook Style
        style.configure("TNotebook", background=C["bg"], borderwidth=0, highlightthickness=0)
        style.configure("TNotebook.Tab", background=C["surface"], foreground=C["text2"], 
                        font=("Segoe UI", 10, "bold"), padding=[16, 6], borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", C["accent"]), ("active", C["surface2"])],
                  foreground=[("selected", C["white"]), ("active", C["text"])])
                  
        # Configure Treeview Style
        style.configure("Treeview", background=C["surface"], foreground=C["text"], 
                        fieldbackground=C["surface"], rowheight=25, borderwidth=0, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background=C["surface2"], foreground=C["text"], 
                        font=("Segoe UI", 10, "bold"), borderwidth=0, relief="flat")
        style.map("Treeview", 
                  background=[("selected", C["accent"])], 
                  foreground=[("selected", C["white"])])


    # ── Window setup ──────────────────────────────────────────────────────────
    def _build_window(self):
        self.title("Railway File Organiser")
        self.configure(bg=C["bg"])
        self.geometry("820x660")
        self.minsize(820, 620)
        self.resizable(True, True)

        # Centre on screen
        self.update_idletasks()
        w, h = 820, 660
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ══════════════════════════════════════════════════════════════════════════
    # UI CONSTRUCTION
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # ── Top header bar ────────────────────────────────────────────────────
        self._build_header()
        # ── Status pill ───────────────────────────────────────────────────────
        self._build_status_bar()
        
        # ── Notebook containing tabs ──────────────────────────────────────────
        self.notebook = ttk.Notebook(self, style="TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        # Tab 1: Dashboard
        self.tab_dashboard = tk.Frame(self.notebook, bg=C["bg"])
        self.notebook.add(self.tab_dashboard, text="  📊 Dashboard  ")
        
        # Tab 2: File Register
        self.tab_register = tk.Frame(self.notebook, bg=C["bg"])
        self.notebook.add(self.tab_register, text="  📂 File Register  ")
        
        # ── Tab 1: Dashboard content (two columns) ───────────────────────────
        content = tk.Frame(self.tab_dashboard, bg=C["bg"])
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        self._build_left_panel(content)
        self._build_right_panel(content)
        
        # ── Tab 2: File Register content ─────────────────────────────────────
        self._build_register_tab(self.tab_register)
        
        # ── Bottom bar ────────────────────────────────────────────────────────
        self._build_bottom_bar()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=C["surface"], pady=0)
        hdr.pack(fill="x")

        # Accent stripe
        tk.Frame(hdr, bg=C["accent"], height=4).pack(fill="x")

        inner = tk.Frame(hdr, bg=C["surface"], padx=24, pady=14)
        inner.pack(fill="x")

        # Icon + Title
        left = tk.Frame(inner, bg=C["surface"])
        left.pack(side="left")

        tk.Label(left, text="  RAILWAY FILE ORGANISER",
                 bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Label(left, text="  Smart. Automatic. Offline.",
                 bg=C["surface"], fg=C["text2"],
                 font=("Segoe UI", 10)).pack(side="left", pady=(6, 0))

        # Version badge
        badge = tk.Label(inner, text=" v1.0 ", bg=C["accent"],
                         fg=C["white"], font=FONT_BADGE, padx=6, pady=2)
        badge.pack(side="right")

    # ── Status bar ────────────────────────────────────────────────────────────
    def _build_status_bar(self):
        bar = tk.Frame(self, bg=C["bg"], pady=10, padx=24)
        bar.pack(fill="x")

        self._status_dot  = tk.Label(bar, text="●", fg=C["text3"],
                                      bg=C["bg"], font=("Segoe UI", 16))
        self._status_dot.pack(side="left")

        self._status_text = tk.Label(bar, text="Not watching",
                                      bg=C["bg"], fg=C["text2"],
                                      font=("Segoe UI", 11, "bold"))
        self._status_text.pack(side="left", padx=8)

        self._watch_folder_lbl = tk.Label(bar, text="",
                                           bg=C["bg"], fg=C["text3"],
                                           font=FONT_SMALL)
        self._watch_folder_lbl.pack(side="left")

        # Counter badges
        right = tk.Frame(bar, bg=C["bg"])
        right.pack(side="right")

        tk.Label(right, text="FILES ORGANISED TODAY",
                 bg=C["bg"], fg=C["text3"], font=FONT_SMALL).pack(side="left")
        self._count_lbl = tk.Label(right, text=" 0 ",
                                    bg=C["accent"], fg=C["white"],
                                    font=("Segoe UI", 10, "bold"),
                                    padx=8, pady=1)
        self._count_lbl.pack(side="left", padx=(6, 0))
        self._today_count = 0

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=20)

    # ── Left panel ────────────────────────────────────────────────────────────
    def _build_left_panel(self, parent):
        frame = tk.Frame(parent, bg=C["bg"])
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=12)
        frame.rowconfigure(1, weight=1)

        # ── Big action buttons ────────────────────────────────────────────────
        btn_card = tk.Frame(frame, bg=C["surface"], padx=20, pady=20,
                             highlightbackground=C["border"], highlightthickness=1)
        btn_card.pack(fill="x")

        tk.Label(btn_card, text="Quick Actions",
                 bg=C["surface"], fg=C["text2"],
                 font=FONT_BADGE).pack(anchor="w", pady=(0, 14))

        # START / STOP WATCHING
        self._watch_btn = self._big_button(
            btn_card,
            icon="▶",
            title="Start Auto-Organiser",
            subtitle="Watches your Downloads folder\nand sorts files automatically",
            color=C["green"],
            color2=C["green2"],
            command=self._toggle_watch
        )
        self._watch_btn.pack(fill="x", pady=(0, 10))

        # ORGANISE NOW (batch)
        self._big_button(
            btn_card,
            icon="⚡",
            title="Organise a Folder Now",
            subtitle="Pick any folder and sort all\nexisting files immediately",
            color=C["accent"],
            color2=C["accent2"],
            command=self._run_batch_now
        ).pack(fill="x", pady=(0, 10))

        # OPEN EXCEL LOG
        self._big_button(
            btn_card,
            icon="📊",
            title="Open File Register",
            subtitle="View all organised files\nwith clickable links",
            color="#7C3AED",
            color2="#6D28D9",
            command=self._open_excel
        ).pack(fill="x", pady=(0, 10))

        # TEST POPUP BUTTON
        self._big_button(
            btn_card,
            icon="🔔",
            title="Test Popup Alert",
            subtitle="Verify the file detection popup\nworks on this machine",
            color="#0E7490",
            color2="#0C687E",
            command=self._test_popup
        ).pack(fill="x")

        # ── Activity log ──────────────────────────────────────────────────────
        tk.Label(frame, text="Recent Activity",
                 bg=C["bg"], fg=C["text2"],
                 font=FONT_BADGE).pack(anchor="w", pady=(18, 6))

        log_frame = tk.Frame(frame, bg=C["surface"],
                              highlightbackground=C["border"], highlightthickness=1)
        log_frame.pack(fill="both", expand=True)

        self._log_text = tk.Text(log_frame, bg=C["surface"], fg=C["text2"],
                                  font=FONT_MONO, relief="flat",
                                  state="disabled", wrap="word",
                                  insertbackground=C["text"],
                                  selectbackground=C["accent"],
                                  padx=12, pady=10)
        self._log_text.pack(fill="both", expand=True)

        # Tag styles
        self._log_text.tag_configure("ok",   foreground="#22C55E")
        self._log_text.tag_configure("info", foreground=C["text2"])
        self._log_text.tag_configure("warn", foreground="#FBBF24")
        self._log_text.tag_configure("file", foreground="#60A5FA", font=("Consolas", 9, "bold"))
        self._log_text.tag_configure("cat",  foreground="#A78BFA")

        self._log("Welcome! Click 'Start Auto-Organiser' to begin.", "info")

    # ── Right panel ───────────────────────────────────────────────────────────
    def _build_right_panel(self, parent):
        frame = tk.Frame(parent, bg=C["bg"])
        frame.grid(row=0, column=1, sticky="nsew", pady=12)

        # ── Folders card ──────────────────────────────────────────────────────
        card = tk.Frame(frame, bg=C["surface"], padx=18, pady=16,
                         highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill="x")

        tk.Label(card, text="Folder Configuration",
                 bg=C["surface"], fg=C["text2"],
                 font=FONT_BADGE).pack(anchor="w", pady=(0, 12))

        self._watch_var = tk.StringVar(value=self.cfg["watch_folder"])
        self._base_var  = tk.StringVar(value=self.cfg["base_folder"])

        self._folder_row(card, "Watch Folder",
                          "New files here get sorted", self._watch_var,
                          lambda: self._pick_folder(self._watch_var))

        tk.Frame(card, bg=C["border"], height=1).pack(fill="x", pady=10)

        self._folder_row(card, "Organised Files Folder",
                          "Files go here after sorting", self._base_var,
                          lambda: self._pick_folder(self._base_var))

        # Save button
        tk.Button(card, text="Save Folders",
                  bg=C["accent"], fg=C["white"],
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=16, pady=6,
                  cursor="hand2",
                  command=self._save_folders).pack(anchor="e", pady=(12, 0))

        # ── How it works card ─────────────────────────────────────────────────
        how = tk.Frame(frame, bg=C["surface"], padx=18, pady=16,
                        highlightbackground=C["border"], highlightthickness=1)
        how.pack(fill="x", pady=(12, 0))

        tk.Label(how, text="How It Works",
                 bg=C["surface"], fg=C["text2"],
                 font=FONT_BADGE).pack(anchor="w", pady=(0, 10))

        steps = [
            ("1", "A file is downloaded",        C["accent"]),
            ("2", "Software detects it",          "#7C3AED"),
            ("3", "Suggests the right folder",    C["yellow"]),
            ("4", "You confirm in 1 click",       C["green"]),
            ("5", "File moved + Excel updated",   "#EC4899"),
        ]
        for num, text, color in steps:
            row = tk.Frame(how, bg=C["surface"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=f" {num} ", bg=color, fg=C["white"],
                     font=("Segoe UI", 9, "bold"),
                     padx=4, pady=1).pack(side="left")
            tk.Label(row, text=f"  {text}", bg=C["surface"],
                     fg=C["text"], font=("Segoe UI", 10)).pack(side="left")

        # ── Categories summary ────────────────────────────────────────────────
        cat_frame = tk.Frame(frame, bg=C["surface"], padx=18, pady=14,
                              highlightbackground=C["border"], highlightthickness=1)
        cat_frame.pack(fill="both", expand=True, pady=(12, 0))

        tk.Label(cat_frame, text="Known Categories",
                 bg=C["surface"], fg=C["text2"],
                 font=FONT_BADGE).pack(anchor="w", pady=(0, 8))

        cats = self.cat.get_all_categories()[:14]
        for i, cat in enumerate(cats):
            color = self.cat.get_category_color(cat)
            row = tk.Frame(cat_frame, bg=C["surface"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text="  ", bg=color, width=2).pack(side="left")
            tk.Label(row, text=f"  {cat}",
                     bg=C["surface"], fg=C["text"],
                     font=("Segoe UI", 9)).pack(side="left")

        rem = len(self.cat.get_all_categories()) - 14
        if rem > 0:
            tk.Label(cat_frame, text=f"  + {rem} more categories…",
                     bg=C["surface"], fg=C["text3"],
                     font=FONT_SMALL).pack(anchor="w", pady=(4, 0))

    # ── Bottom bar ────────────────────────────────────────────────────────────
    def _build_bottom_bar(self):
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=20)
        bar = tk.Frame(self, bg=C["surface"], pady=10, padx=20)
        bar.pack(fill="x")

        tk.Label(bar, text="Railway File Organiser  |  100% Offline  |  No internet required",
                 bg=C["surface"], fg=C["text3"],
                 font=FONT_SMALL).pack(side="left")

        tk.Button(bar, text="Open Log Folder",
                  bg=C["surface"], fg=C["text2"],
                  font=FONT_SMALL, relief="flat",
                  cursor="hand2",
                  command=self._open_base_folder).pack(side="right")

    def _build_register_tab(self, parent):
        """Build the File Register tab with embedded Excel viewer & verification."""
        frame = tk.Frame(parent, bg=C["bg"], pady=10)
        frame.pack(fill="both", expand=True)
        
        # ── Summary / Stat Card ───────────────────────────────────────────────
        self._reg_stat_card = tk.Frame(frame, bg=C["surface"], padx=16, pady=10,
                                       highlightbackground=C["border"], highlightthickness=1)
        self._reg_stat_card.pack(fill="x", pady=(0, 10))
        
        self._reg_stat_lbl = tk.Label(self._reg_stat_card, 
                                      text="Loading register data...",
                                      bg=C["surface"], fg=C["text2"],
                                      font=("Segoe UI", 10, "bold"))
        self._reg_stat_lbl.pack(side="left")
        
        # Warning label if some files are missing
        self._reg_warn_lbl = tk.Label(self._reg_stat_card,
                                      text="",
                                      bg=C["surface"], fg=C["red"],
                                      font=("Segoe UI", 10, "bold"))
        self._reg_warn_lbl.pack(side="right")
        
        # ── Treeview Table ────────────────────────────────────────────────────
        table_frame = tk.Frame(frame, bg=C["bg"])
        table_frame.pack(fill="both", expand=True)
        
        columns = ("sno", "datetime", "original", "category", "final", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Columns config
        self.tree.heading("sno", text="S.No")
        self.tree.heading("datetime", text="Date & Time")
        self.tree.heading("original", text="Original File Name")
        self.tree.heading("category", text="Category")
        self.tree.heading("final", text="Final File Name")
        self.tree.heading("status", text="Status")
        
        self.tree.column("sno", width=50, anchor="center")
        self.tree.column("datetime", width=140, anchor="w")
        self.tree.column("original", width=220, anchor="w")
        self.tree.column("category", width=140, anchor="w")
        self.tree.column("final", width=220, anchor="w")
        self.tree.column("status", width=100, anchor="center")
        
        # Using grid to align scrollbars with treeview
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Tag colors for Treeview rows
        self.tree.tag_configure("missing", foreground="#F87171") # light red
        self.tree.tag_configure("existing", foreground="#34D399") # light green
        
        # Bind double click on row to open the file
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        
        # ── Buttons Bar ───────────────────────────────────────────────────────
        btn_bar = tk.Frame(frame, bg=C["bg"], pady=12)
        btn_bar.pack(fill="x")
        
        # Refresh / Verify Button
        self.btn_refresh = tk.Button(btn_bar, text="🔄 Refresh & Verify Files",
                                     bg=C["accent"], fg=C["white"],
                                     font=("Segoe UI", 10, "bold"),
                                     relief="flat", padx=14, pady=6, cursor="hand2",
                                     command=self._load_and_verify_register)
        self.btn_refresh.pack(side="left", padx=(0, 10))
        
        # Rebuild Log Button
        self.btn_rebuild = tk.Button(btn_bar, text="🛠️ Rebuild Excel Log from Folder",
                                     bg=C["yellow"], fg=C["bg"],
                                     font=("Segoe UI", 10, "bold"),
                                     relief="flat", padx=14, pady=6, cursor="hand2",
                                     command=self._confirm_and_rebuild_log)
        self.btn_rebuild.pack(side="left", padx=(0, 10))
        
        # Open Excel Log Button
        self.btn_open_excel = tk.Button(btn_bar, text="📊 Open Excel File",
                                        bg="#7C3AED", fg=C["white"],
                                        font=("Segoe UI", 10, "bold"),
                                        relief="flat", padx=14, pady=6, cursor="hand2",
                                        command=self._open_excel)
        self.btn_open_excel.pack(side="left", padx=(0, 10))

        # Open Base Folder Button
        self.btn_open_folder = tk.Button(btn_bar, text="📁 Open Organised Folder",
                                         bg=C["surface2"], fg=C["text"],
                                         font=("Segoe UI", 10, "bold"),
                                         relief="flat", padx=14, pady=6, cursor="hand2",
                                         command=self._open_base_folder)
        self.btn_open_folder.pack(side="left")

    def _on_tree_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        file_path = self._row_paths.get(item_id)
        if file_path and os.path.exists(file_path):
            try:
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file:\n{e}")
        else:
            messagebox.showerror("File Missing", "This file does not exist on disk!")

    def _load_and_verify_register(self):
        """Read the Excel log, verify file existence, and populate the Treeview (threaded)."""
        self._reg_stat_lbl.config(text="Reading log and verifying files on disk...")
        self._reg_warn_lbl.config(text="")
        
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_paths = {}
        
        # Run loading in background thread
        threading.Thread(target=self._verify_worker, daemon=True).start()

    def _verify_worker(self):
        logger = ExcelLogger(self.cfg.log_path(), open_after_update=False)
        entries = logger.read_log()
        
        results = []
        total = len(entries)
        existing = 0
        missing = 0
        
        for entry in entries:
            path = entry["path"]
            exists = os.path.exists(path) if path else False
            if exists:
                existing += 1
                status = "🟢 Existing"
                tag = "existing"
            else:
                missing += 1
                status = "🔴 Missing"
                tag = "missing"
                
            results.append((entry, status, tag))
            
        # Put results into UI queue to display safely on main thread
        self._ui_queue.put(("register_data", {
            "results": results,
            "total": total,
            "existing": existing,
            "missing": missing
        }))

    def _display_register_data(self, data):
        # Clear tree again to prevent duplicate additions
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_paths = {}
        
        results = data["results"]
        for entry, status, tag in results:
            item_id = self.tree.insert("", "end", values=(
                entry["sno"],
                entry["datetime"],
                entry["original"],
                entry["category"],
                entry["final"],
                status
            ), tags=(tag,))
            self._row_paths[item_id] = entry["path"]
            
        # Update labels
        self._reg_stat_lbl.config(text=f"Total Organised: {data['total']}   |   🟢 Verified: {data['existing']}   |   🔴 Missing: {data['missing']}")
        if data["missing"] > 0:
            self._reg_warn_lbl.config(text=f"⚠️ WARNING: {data['missing']} file(s) missing from folders! ")
        else:
            self._reg_warn_lbl.config(text="✅ All files successfully verified! ")

    def _confirm_and_rebuild_log(self):
        """Ask for user permission before rewriting the Excel log."""
        msg = ("This will scan your organised folder:\n"
               f"'{self.cfg['base_folder']}'\n\n"
               "It will rebuild a fresh Excel Log from scratch based on the files physically present in the folders.\n"
               "Existing timestamps and original file names will be preserved where possible.\n\n"
               "Would you like to proceed with rebuilding the Excel log?")
               
        if not messagebox.askyesno("Confirm Rebuild Excel Log", msg):
            return
            
        self._reg_stat_lbl.config(text="Rebuilding Excel Log...")
        self._reg_warn_lbl.config(text="")
        
        # Clear items
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        threading.Thread(target=self._rebuild_worker, daemon=True).start()

    def _rebuild_worker(self):
        base_folder = Path(self.cfg["base_folder"])
        log_path = Path(self.cfg.log_path())
        
        if not base_folder.exists():
            self._ui_queue.put(("warn", "Organised folder does not exist."))
            self._ui_queue.put(("show_error", f"Organised folder not found:\n{base_folder}"))
            return
            
        # 1. Read existing metadata if Excel log exists
        metadata = {}
        logger = ExcelLogger(str(log_path), open_after_update=False)
        if log_path.exists():
            existing_entries = logger.read_log()
            for entry in existing_entries:
                p = entry.get("path")
                if p:
                    # Normalize path to match easily
                    norm_p = os.path.normpath(p).lower()
                    metadata[norm_p] = {
                        "original": entry.get("original"),
                        "datetime": entry.get("datetime"),
                        "category": entry.get("category")
                    }
                    
        # 2. Scan folders
        found_entries = []
        for root, dirs, files in os.walk(str(base_folder)):
            for file in files:
                file_path = os.path.join(root, file)
                # Ignore Excel log itself
                if os.path.normpath(file_path) == os.path.normpath(str(log_path)):
                    continue
                # Ignore hidden files, system files, temp files
                if file.startswith("~") or file.startswith("."):
                    continue
                    
                # Determine category
                try:
                    rel_parts = Path(file_path).relative_to(base_folder).parts
                    if len(rel_parts) > 1:
                        category = rel_parts[0]
                    else:
                        category = "MISCELLANEOUS FILES"
                except Exception:
                    category = "MISCELLANEOUS FILES"
                    
                norm_p = os.path.normpath(file_path).lower()
                
                # Check metadata map
                if norm_p in metadata:
                    orig = metadata[norm_p]["original"]
                    dt = metadata[norm_p]["datetime"]
                else:
                    orig = file
                    try:
                        mtime = os.path.getmtime(file_path)
                        dt = datetime.fromtimestamp(mtime).strftime("%d-%b-%Y  %H:%M:%S")
                    except Exception:
                        dt = datetime.now().strftime("%d-%b-%Y  %H:%M:%S")
                        
                found_entries.append({
                    "original": orig,
                    "final": file,
                    "category": category,
                    "path": file_path,
                    "datetime": dt
                })
                
        # Sort found_entries by datetime
        def get_sort_key(entry):
            try:
                return datetime.strptime(entry["datetime"], "%d-%b-%Y  %H:%M:%S")
            except Exception:
                return datetime.min
                
        found_entries.sort(key=get_sort_key)
        
        # 3. Overwrite log file
        res = logger.write_all_entries(found_entries)
        
        if res["success"]:
            self._ui_queue.put(("ok", f"Excel log rebuilt successfully with {len(found_entries)} files."))
            # Reload tree
            self._load_and_verify_register()
            # Inform user
            self._ui_queue.put(("rebuild_success", len(found_entries)))
        else:
            self._ui_queue.put(("warn", f"Failed to rebuild Excel log: {res['error']}"))
            self._ui_queue.put(("show_error", f"Could not rebuild Excel log:\n{res['error']}"))


    # ══════════════════════════════════════════════════════════════════════════
    # WIDGET HELPERS
    # ══════════════════════════════════════════════════════════════════════════
    def _big_button(self, parent, icon, title, subtitle, color, color2, command):
        """Large action button with icon + title + subtitle."""
        btn = tk.Frame(parent, bg=color, cursor="hand2",
                        highlightbackground=color2, highlightthickness=1)

        inner = tk.Frame(btn, bg=color, padx=14, pady=12)
        inner.pack(fill="both")

        left = tk.Frame(inner, bg=color)
        left.pack(side="left", fill="y")

        tk.Label(left, text=icon, bg=color, fg=C["white"],
                 font=("Segoe UI", 22)).pack()

        right = tk.Frame(inner, bg=color)
        right.pack(side="left", padx=12)

        tk.Label(right, text=title, bg=color, fg=C["white"],
                 font=("Segoe UI", 12, "bold"), anchor="w").pack(anchor="w")
        tk.Label(right, text=subtitle, bg=color, fg="#CCDDFF",
                 font=("Segoe UI", 9), anchor="w", justify="left").pack(anchor="w")

        # Make entire button clickable
        for widget in [btn, inner, left, right] + inner.winfo_children() + right.winfo_children():
            widget.bind("<Button-1>", lambda e: command())
            widget.bind("<Enter>",    lambda e, b=btn: b.config(bg=color2))
            widget.bind("<Leave>",    lambda e, b=btn: b.config(bg=color))

        return btn

    def _folder_row(self, parent, label, hint, var, browse_cmd):
        tk.Label(parent, text=label, bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(parent, text=hint, bg=C["surface"], fg=C["text3"],
                 font=FONT_SMALL).pack(anchor="w", pady=(0, 4))

        row = tk.Frame(parent, bg=C["surface"])
        row.pack(fill="x")

        entry = tk.Entry(row, textvariable=var,
                         bg=C["surface2"], fg=C["text"],
                         insertbackground=C["text"],
                         relief="flat", font=("Segoe UI", 9),
                         disabledbackground=C["surface2"],
                         disabledforeground=C["text2"])
        entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))

        tk.Button(row, text="Browse",
                  bg=C["border"], fg=C["text"],
                  font=FONT_SMALL, relief="flat",
                  padx=10, pady=4, cursor="hand2",
                  command=browse_cmd).pack(side="left")

    # ══════════════════════════════════════════════════════════════════════════
    # ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _auto_start_if_configured(self):
        """Auto-start watching on launch — always on by default."""
        if not self.watching:
            watch_folder = self.cfg["watch_folder"]
            if os.path.isdir(watch_folder):
                self._start_watching()
                self._log("Auto-Organiser started automatically.", "ok")
            else:
                self._log("Auto-start skipped: watch folder not found. Please set folders first.", "warn")

    def _toggle_watch(self):
        if self.watching:
            self._stop_watching()
        else:
            self._start_watching()

    def _test_popup(self):
        """Show a test popup to verify the alert system works on this machine."""
        self._log("Testing popup system...", "info")
        # Step 1: try simple messagebox first
        try:
            result = messagebox.askquestion(
                "POPUP TEST - Step 1",
                "If you can read this, basic dialogs work!\n\n"
                "Click YES to also test the full file-detection popup.\n"
                "Click NO to cancel.",
                icon="question"
            )
            if result != "yes":
                self._log("Test cancelled by user.", "warn")
                return
        except Exception as e:
            self._log(f"ERROR: Even messagebox failed: {e}", "warn")
            return

        # Step 2: try the full custom popup
        import tempfile
        try:
            tmp = Path(tempfile.gettempdir()) / "TEST_Railway_popup_demo.pdf"
            tmp.write_text("test", encoding="utf-8")
            fake_result = {"category": "TEST CATEGORY", "confidence": 0.99}
            self._show_confirm_dialog(str(tmp), tmp.name, fake_result)
            self._log("Full popup launched - check if it appeared!", "info")
        except Exception as e:
            import traceback
            self._log(f"POPUP ERROR: {e}", "warn")
            messagebox.showerror("Popup Failed",
                f"The full popup failed with error:\n{e}\n\n"
                f"Basic messageboxes work. The Toplevel popup has a bug.\n\n"
                f"{traceback.format_exc()}")

    def _start_watching(self):
        folder = self.cfg["watch_folder"]
        if not os.path.isdir(folder):
            messagebox.showerror("Folder Not Found",
                f"Watch folder not found:\n{folder}\n\nPlease update it in the settings.")
            return

        try:
            self.watcher = FileWatcher(
                folder,
                on_file_ready=lambda p: self._ui_queue.put(("new_file", p)),
                stable_wait=float(self.cfg["stable_wait_seconds"]),
                stable_count=2
            )
            self.watcher.start()
            self.watching = True
            self._refresh_status()
            self._log(f"Now watching: {folder}", "ok")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _stop_watching(self):
        if self.watcher:
            self.watcher.stop()
            self.watcher = None
        self.watching = False
        self._refresh_status()
        self._log("Stopped watching.", "warn")

    def _run_batch_now(self):
        """Let user pick a folder and batch-organise all files in it."""
        src = filedialog.askdirectory(
            title="Select the folder to organise",
            initialdir=self.cfg["watch_folder"]
        )
        if not src:
            return
        threading.Thread(target=self._batch_worker, args=(src,), daemon=True).start()

    def _batch_worker(self, src_folder: str):
        SKIP = {".crdownload", ".part", ".tmp", ".download"}
        mover  = FileMover(self.cfg["base_folder"])
        logger = ExcelLogger(self.cfg.log_path(),
                             open_after_update=False)

        files = [f for f in os.listdir(src_folder)
                 if os.path.isfile(os.path.join(src_folder, f))
                 and Path(f).suffix.lower() not in SKIP]

        if not files:
            self._ui_queue.put(("info", "No files found in selected folder."))
            return

        self._ui_queue.put(("info", f"Batch: found {len(files)} files..."))

        # Build map
        file_map = []
        for fname in files:
            src = os.path.join(src_folder, fname)
            result = self.cat.categorise(src)
            file_map.append({
                "filename": fname,
                "src_path": src,
                "category": result["category"],
                "confidence": result["confidence"]
            })

        # Show batch preview dialog on main thread
        self._ui_queue.put(("batch_preview", file_map))

    def _open_excel(self):
        log = self.cfg.log_path()
        if os.path.exists(log):
            os.startfile(log)
        else:
            messagebox.showinfo("No Log Yet",
                "No files have been organised yet.\n"
                "The Excel register will be created automatically\n"
                "after the first file is organised.")

    def _open_base_folder(self):
        base = self.cfg["base_folder"]
        Path(base).mkdir(parents=True, exist_ok=True)
        os.startfile(base)

    def _pick_folder(self, var: tk.StringVar):
        chosen = filedialog.askdirectory(initialdir=var.get() or str(Path.home()))
        if chosen:
            var.set(chosen)

    def _save_folders(self):
        self.cfg["watch_folder"] = self._watch_var.get().strip()
        self.cfg["base_folder"]  = self._base_var.get().strip()
        self.cfg.save()
        Path(self.cfg["base_folder"]).mkdir(parents=True, exist_ok=True)
        self._refresh_status()
        self._log("Folders saved.", "ok")
        messagebox.showinfo("Saved", "Folder settings saved successfully!")

    # ══════════════════════════════════════════════════════════════════════════
    # FILE PROCESSING PIPELINE
    # ══════════════════════════════════════════════════════════════════════════
    def _process_file(self, filepath: str):
        """Called when a new file is ready. Shows confirm dialog."""
        try:
            filename = Path(filepath).name
            self._log(f"Processing: {filename}", "info")
            result = self.cat.categorise(filepath)
            self._show_confirm_dialog(filepath, filename, result)
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self._log(f"ERROR showing popup for {Path(filepath).name}: {e}", "warn")
            logging.error(f"_process_file failed: {err}")
            try:
                messagebox.showerror("File Organiser - Error",
                    f"Could not show popup for:\n{Path(filepath).name}\n\nError: {e}")
            except Exception:
                pass

    def _show_confirm_dialog(self, filepath, filename, cat_result):
        """Unmissable popup — forces itself to the front on Windows 10/11."""
        dest_folder = str(Path(self.cfg["base_folder"]) / cat_result["category"])
        is_test = filename.startswith("TEST_Railway_popup_demo")

        dlg = tk.Toplevel(self)
        dlg.title("NEW FILE DETECTED - Railway File Organiser")
        dlg.configure(bg=C["bg"])
        dlg.resizable(False, False)

        w, h = 640, 520
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        # ── Force to front using Windows API ─────────────────────────────────
        def _force_to_front():
            try:
                dlg.update_idletasks()
                hwnd = ctypes.windll.user32.FindWindowW(None, dlg.title())
                if not hwnd:
                    hwnd = int(dlg.frame(), 16)
                _force_window_front(hwnd)
                _flash_window(hwnd)
                dlg.attributes("-topmost", True)
                dlg.lift()
                dlg.focus_force()
            except Exception:
                pass

        # Try to force focus immediately, then again after 300ms and 800ms
        dlg.after(50,  _force_to_front)
        dlg.after(300, _force_to_front)
        dlg.after(800, _force_to_front)

        # Play system alert sound
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            dlg.bell()

        # ── Pulsing top stripe ────────────────────────────────────────────────
        stripe_colors = [C["yellow"], "#F59E0B", C["yellow"]]
        stripe = tk.Frame(dlg, bg=C["yellow"], height=8)
        stripe.pack(fill="x")

        def _pulse_stripe(i=0):
            if dlg.winfo_exists():
                stripe.config(bg=stripe_colors[i % len(stripe_colors)])
                dlg.after(500, _pulse_stripe, i + 1)
        _pulse_stripe()

        # ── Header ───────────────────────────────────────────────────────────
        hdr = tk.Frame(dlg, bg=C["bg"], padx=24, pady=14)
        hdr.pack(fill="x")
        title_text = "🧪 TEST — POPUP IS WORKING!" if is_test else "🔔  NEW FILE DETECTED"
        title_color = C["green"] if is_test else C["yellow"]
        tk.Label(hdr, text=title_text,
                 bg=C["bg"], fg=title_color,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w")
        tk.Label(hdr,
                 text="Review below and approve or skip. Nothing happens until you click YES.",
                 bg=C["bg"], fg=C["text2"],
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        tk.Frame(dlg, bg=C["border"], height=1).pack(fill="x", padx=20)

        # ── File card ────────────────────────────────────────────────────────
        card = tk.Frame(dlg, bg=C["surface"], padx=20, pady=14,
                         highlightbackground=title_color, highlightthickness=2)
        card.pack(fill="x", padx=20, pady=14)

        tk.Label(card, text="FILE DETECTED",
                 bg=C["surface"], fg=C["text3"], font=FONT_BADGE).pack(anchor="w")
        display = (filename[:70]+"...") if len(filename) > 73 else filename
        tk.Label(card, text=display,
                 bg=C["surface"], fg="#FBBF24",
                 font=("Segoe UI", 11, "bold"),
                 wraplength=580, justify="left").pack(anchor="w", pady=(3, 12))

        tk.Label(card, text="WILL BE COPIED TO",
                 bg=C["surface"], fg=C["text3"], font=FONT_BADGE).pack(anchor="w")

        path_frame = tk.Frame(card, bg=C["surface2"], padx=10, pady=8)
        path_frame.pack(fill="x", pady=(4, 0))
        path_lbl = tk.Label(path_frame, text=dest_folder,
                            bg=C["surface2"], fg="#60A5FA",
                            font=("Consolas", 9),
                            wraplength=580, justify="left")
        path_lbl.pack(anchor="w")

        # ── Category selector ────────────────────────────────────────────────
        sel = tk.Frame(dlg, bg=C["bg"], padx=20)
        sel.pack(fill="x")
        tk.Label(sel, text="Change category (optional):",
                 bg=C["bg"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        cat_var = tk.StringVar(value=cat_result["category"])
        all_cats = self.cat.get_all_categories()
        combo = ttk.Combobox(sel, textvariable=cat_var, values=all_cats,
                              font=("Segoe UI", 10), width=44, state="normal")
        combo.pack(anchor="w")

        def update_path(*_):
            new_dest = str(Path(self.cfg["base_folder"]) / (cat_var.get().strip() or cat_result["category"]))
            path_lbl.config(text=new_dest)
        cat_var.trace_add("write", update_path)

        tk.Label(sel,
                 text="✅  The original file will NOT be deleted. Only a copy is made.",
                 bg=C["bg"], fg=C["green"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))

        # ── Countdown label ───────────────────────────────────────────────────
        countdown_var = tk.StringVar(value="")
        countdown_lbl = tk.Label(sel, textvariable=countdown_var,
                                  bg=C["bg"], fg=C["text3"],
                                  font=("Segoe UI", 9))
        countdown_lbl.pack(anchor="w", pady=(4, 0))

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_frame = tk.Frame(dlg, bg=C["bg"], padx=20, pady=16)
        btn_frame.pack(fill="x")

        def do_confirm():
            chosen = cat_var.get().strip() or cat_result["category"]
            final_dest = str(Path(self.cfg["base_folder"]) / chosen)
            dlg.destroy()
            if not is_test:
                threading.Thread(
                    target=self._copy_and_log,
                    args=(filepath, filename, chosen, final_dest),
                    daemon=True
                ).start()
            else:
                self._log("TEST POPUP: Confirmed successfully! Popup is working.", "ok")

        def do_skip():
            dlg.destroy()
            if not is_test:
                self._log(f"SKIPPED: {filename}", "warn")
            else:
                self._log("TEST POPUP: Skipped. Popup is working correctly.", "ok")

        # Big YES button
        yes_btn = tk.Button(btn_frame,
                  text="  ✅  YES — Copy & Organise  ",
                  bg=C["green"], fg="#FFFFFF",
                  font=("Segoe UI", 13, "bold"),
                  relief="flat", padx=24, pady=13,
                  cursor="hand2",
                  command=do_confirm)
        yes_btn.pack(side="left", padx=(0, 14))

        # Hover effect on YES
        yes_btn.bind("<Enter>", lambda e: yes_btn.config(bg="#15803D"))
        yes_btn.bind("<Leave>", lambda e: yes_btn.config(bg=C["green"]))

        tk.Button(btn_frame,
                  text="NO — Skip",
                  bg=C["red"], fg=C["white"],
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=18, pady=13,
                  cursor="hand2",
                  command=do_skip).pack(side="left")

        tk.Label(btn_frame,
                 text="Closing this window = Skip",
                 bg=C["bg"], fg=C["text3"], font=FONT_SMALL
                 ).pack(side="right", pady=(8, 0))

        # ── Auto-dismiss countdown (90 seconds) ───────────────────────────────
        timeout_secs = [90]
        def _tick():
            if not dlg.winfo_exists():
                return
            s = timeout_secs[0]
            countdown_var.set(f"Auto-skip in {s}s if no action taken")
            if s <= 0:
                do_skip()
                return
            timeout_secs[0] -= 1
            dlg.after(1000, _tick)
        _tick()

        dlg.protocol("WM_DELETE_WINDOW", do_skip)
        dlg.grab_set()

    def _copy_and_log(self, filepath, filename, category, dest_folder_path):
        """Copy file and log. Shows clear success OR error popup after."""
        mover  = FileMover(self.cfg["base_folder"])
        logger = ExcelLogger(self.cfg.log_path(), open_after_update=False)
        result = mover.move(filepath, category)

        if result["success"]:
            logger.log(result["original_filename"],
                       result["final_filename"],
                       category,
                       result["destination"])
            self._ui_queue.put(("ok",    f"{filename}  copied to  {category}"))
            self._ui_queue.put(("count", 1))
            self._ui_queue.put(("refresh_register", None))
            # Show clear success popup on main thread
            self._ui_queue.put(("show_success", {
                "filename":    result["final_filename"],
                "category":    category,
                "destination": result["destination"],
                "log_path":    self.cfg.log_path()
            }))
        else:
            self._ui_queue.put(("warn", f"FAILED to copy {filename}: {result['error']}"))
            self._ui_queue.put(("show_error", f"Could not copy file:\n{filename}\n\nReason: {result['error']}"))

    # Keep old name as alias so batch still works
    def _move_and_log(self, filepath, filename, category):
        dest = str(Path(self.cfg["base_folder"]) / category)
        self._copy_and_log(filepath, filename, category, dest)


    def _do_batch_move(self, file_map):
        """Execute batch move after preview confirmation."""
        mover  = FileMover(self.cfg["base_folder"])
        logger = ExcelLogger(self.cfg.log_path(), open_after_update=False)
        ok = fail = 0

        for item in file_map:
            r = mover.move(item["src_path"], item["category"])
            if r["success"]:
                logger.log(r["original_filename"], r["final_filename"],
                           item["category"], r["destination"])
                self._ui_queue.put(("ok", f"{item['filename'][:45]}  ->  {item['category']}"))
                ok += 1
            else:
                self._ui_queue.put(("warn", f"Failed: {item['filename']}"))
                fail += 1
            self._ui_queue.put(("count", 1))

        self._ui_queue.put(("info", f"Done — {ok} files organised, {fail} failed."))
        self._ui_queue.put(("refresh_register", None))
        if os.path.exists(self.cfg.log_path()):
            os.startfile(self.cfg.log_path())

    # ══════════════════════════════════════════════════════════════════════════
    # BATCH PREVIEW WINDOW
    # ══════════════════════════════════════════════════════════════════════════
    def _show_batch_preview(self, file_map):
        dlg = tk.Toplevel(self)
        dlg.title("Batch Organiser — Preview")
        dlg.configure(bg=C["bg"])
        dlg.geometry("960x580")
        dlg.grab_set()

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        dlg.geometry(f"960x580+{(sw-960)//2}+{(sh-580)//2}")

        tk.Frame(dlg, bg=C["accent"], height=5).pack(fill="x")

        hdr = tk.Frame(dlg, bg=C["bg"], padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Preview — {len(file_map)} files will be organised",
                 bg=C["bg"], fg=C["text"],
                 font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Label(hdr, text="  Review and change any category below",
                 bg=C["bg"], fg=C["text2"],
                 font=("Segoe UI", 10)).pack(side="left")

        tk.Frame(dlg, bg=C["border"], height=1).pack(fill="x", padx=20)

        # Table
        frame = tk.Frame(dlg, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=20, pady=10)

        canvas = tk.Canvas(frame, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        sf = tk.Frame(canvas, bg=C["bg"])
        sf.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sf, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        def _on_wheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                pass
        canvas.bind_all("<MouseWheel>", _on_wheel)

        # Headers
        for col, (txt, w) in enumerate([("#", 4), ("File Name", 52), ("Category", 35)]):
            tk.Label(sf, text=txt, bg=C["surface2"], fg=C["text2"],
                     font=FONT_BADGE, width=w, anchor="w",
                     padx=8, pady=5).grid(row=0, column=col, sticky="ew", padx=1, pady=1)

        all_cats = self.cat.get_all_categories()
        cat_vars = []

        for i, item in enumerate(file_map, 1):
            bg = C["surface"] if i % 2 == 0 else C["surface2"]
            tk.Label(sf, text=str(i), bg=bg, fg=C["text3"],
                     font=FONT_MONO, width=4, anchor="center"
                     ).grid(row=i, column=0, sticky="ew", padx=1, pady=1)

            dn = item["filename"]
            if len(dn) > 60: dn = dn[:57] + "..."
            tk.Label(sf, text=dn, bg=bg, fg=C["text"],
                     font=FONT_MONO, width=52, anchor="w", padx=8
                     ).grid(row=i, column=1, sticky="ew", padx=1, pady=1)

            v = tk.StringVar(value=item["category"])
            cat_vars.append(v)
            ttk.Combobox(sf, textvariable=v, values=all_cats,
                          font=("Segoe UI", 9), width=33, state="normal"
                          ).grid(row=i, column=2, sticky="ew", padx=1, pady=1)

        # Buttons
        btns = tk.Frame(dlg, bg=C["bg"], pady=14)
        btns.pack(fill="x", padx=20)

        def _cleanup():
            try:
                canvas.unbind_all("<MouseWheel>")
            except tk.TclError:
                pass

        def proceed():
            for i, item in enumerate(file_map):
                item["category"] = cat_vars[i].get().strip() or item["category"]
            _cleanup()
            dlg.destroy()
            threading.Thread(target=self._do_batch_move,
                             args=(file_map,), daemon=True).start()

        tk.Button(btns, text=f"Organise {len(file_map)} Files Now",
                  bg=C["green"], fg=C["bg"],
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=20, pady=9,
                  cursor="hand2",
                  command=proceed).pack(side="left", padx=(0, 10))

        tk.Button(btns, text="Cancel",
                  bg=C["surface"], fg=C["text2"],
                  font=FONT_BODY, relief="flat",
                  padx=14, pady=9, cursor="hand2",
                  command=lambda: (_cleanup(), dlg.destroy())).pack(side="left")

        dlg.protocol("WM_DELETE_WINDOW", lambda: (_cleanup(), dlg.destroy()))

    # ══════════════════════════════════════════════════════════════════════════
    # UI UPDATE HELPERS
    # ══════════════════════════════════════════════════════════════════════════
    def _log(self, text: str, tag: str = "info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_text.configure(state="normal")
        self._log_text.insert("end", f"[{ts}]  ", "info")
        self._log_text.insert("end", text + "\n", tag)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    def _refresh_status(self):
        if self.watching:
            self._status_dot.config(fg="#22C55E")
            self._status_text.config(text="Watching for new files", fg="#22C55E")
            self._watch_folder_lbl.config(
                text=f"  {self.cfg['watch_folder']}")
            # Update button
            for child in self._watch_btn.winfo_children():
                for c in child.winfo_children():
                    if isinstance(c, tk.Label):
                        if c.cget("font") == str(("Segoe UI", 12, "bold")):
                            c.config(text="Stop Auto-Organiser")
        else:
            self._status_dot.config(fg=C["text3"])
            self._status_text.config(text="Not watching", fg=C["text2"])
            self._watch_folder_lbl.config(text="")
            for child in self._watch_btn.winfo_children():
                for c in child.winfo_children():
                    if isinstance(c, tk.Label):
                        if c.cget("font") == str(("Segoe UI", 12, "bold")):
                            c.config(text="Start Auto-Organiser")

    def _poll_ui_queue(self):
        """Poll the thread-safe queue and update UI from main thread."""
        try:
            while True:
                try:
                    msg = self._ui_queue.get_nowait()
                except queue.Empty:
                    break

                kind, data = msg
                try:
                    if kind == "new_file":
                        self._log(f"Detected: {Path(data).name}", "file")
                        self._process_file(data)
                    elif kind == "ok":
                        self._log(data, "ok")
                    elif kind == "warn":
                        self._log(data, "warn")
                    elif kind == "info":
                        self._log(data, "info")
                    elif kind == "count":
                        self._today_count += data
                        self._count_lbl.config(text=f" {self._today_count} ")
                    elif kind == "batch_preview":
                        self._show_batch_preview(data)
                    elif kind == "show_success":
                        self._show_success_popup(data)
                    elif kind == "show_error":
                        self._show_error_popup(data)
                    elif kind == "register_data":
                        self._display_register_data(data)
                    elif kind == "rebuild_success":
                        messagebox.showinfo("Excel Log Rebuilt",
                                            f"Excel File Register rebuilt!\n"
                                            f"Found {data} file(s).")
                    elif kind == "refresh_register":
                        self._load_and_verify_register()
                except Exception as handler_err:
                    import traceback
                    err_msg = traceback.format_exc()
                    self._log(f"UI ERROR [{kind}]: {handler_err}", "warn")
                    logging.error(f"UI queue handler error: {err_msg}")
        except Exception as poll_err:
            logging.error(f"Poll loop crashed: {poll_err}")
        finally:
            # ALWAYS reschedule — even after a crash
            self.after(200, self._poll_ui_queue)

    def _on_close(self):
        self._stop_watching()
        self.destroy()

    def _show_success_popup(self, info: dict):
        """Clear success popup showing exactly where the file was copied."""
        dlg = tk.Toplevel(self)
        dlg.title("File Organised Successfully")
        dlg.configure(bg=C["bg"])
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        dlg.lift()

        w, h = 580, 340
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        tk.Frame(dlg, bg=C["green"], height=6).pack(fill="x")

        hdr = tk.Frame(dlg, bg=C["bg"], padx=24, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  FILE ORGANISED SUCCESSFULLY",
                 bg=C["bg"], fg=C["green"],
                 font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(hdr, text="Your file has been copied to the folder below. Original is untouched.",
                 bg=C["bg"], fg=C["text2"],
                 font=("Segoe UI", 10)).pack(anchor="w")

        tk.Frame(dlg, bg=C["border"], height=1).pack(fill="x", padx=20)

        card = tk.Frame(dlg, bg=C["surface"], padx=20, pady=14,
                         highlightbackground=C["green"], highlightthickness=1)
        card.pack(fill="x", padx=20, pady=14)

        tk.Label(card, text="FILE", bg=C["surface"],
                 fg=C["text3"], font=FONT_BADGE).pack(anchor="w")
        tk.Label(card, text=info["filename"],
                 bg=C["surface"], fg="#FBBF24",
                 font=("Segoe UI", 10, "bold"),
                 wraplength=520).pack(anchor="w", pady=(2, 10))

        tk.Label(card, text="COPIED TO", bg=C["surface"],
                 fg=C["text3"], font=FONT_BADGE).pack(anchor="w")
        dest_frame = tk.Frame(card, bg=C["surface2"], padx=10, pady=8)
        dest_frame.pack(fill="x", pady=(4, 0))
        tk.Label(dest_frame, text=info["destination"],
                 bg=C["surface2"], fg="#60A5FA",
                 font=("Consolas", 9),
                 wraplength=520, justify="left").pack(anchor="w")

        btn_row = tk.Frame(dlg, bg=C["bg"], padx=20, pady=16)
        btn_row.pack(fill="x")

        def open_folder():
            try:
                os.startfile(str(Path(info["destination"]).parent))
            except Exception:
                pass

        def open_excel():
            if os.path.exists(info["log_path"]):
                os.startfile(info["log_path"])

        tk.Button(btn_row, text="Open Destination Folder",
                  bg=C["accent"], fg=C["white"],
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=16, pady=8,
                  cursor="hand2",
                  command=open_folder).pack(side="left", padx=(0, 10))

        tk.Button(btn_row, text="Open Excel Log",
                  bg="#7C3AED", fg=C["white"],
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=8,
                  cursor="hand2",
                  command=open_excel).pack(side="left", padx=(0, 10))

        tk.Button(btn_row, text="OK",
                  bg=C["surface"], fg=C["text2"],
                  font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=8,
                  cursor="hand2",
                  command=dlg.destroy).pack(side="left")

        # Auto-close after 30s
        dlg.after(30000, lambda: dlg.destroy() if dlg.winfo_exists() else None)

    def _show_error_popup(self, message: str):
        messagebox.showerror("Railway File Organiser — ERROR", message)



# ══════════════════════════════════════════════════════════════════════════════
# FIRST-RUN SETUP — Toplevel on same root. Only ONE tk.Tk() ever created!
# ══════════════════════════════════════════════════════════════════════════════
def first_run(parent_app, cfg):
    """
    CRITICAL FIX: Uses tk.Toplevel instead of tk.Tk().
    Creating a second tk.Tk() breaks Tkinter internals and silently kills
    ALL subsequent Toplevel popups. There must be exactly ONE tk.Tk().
    """
    if os.path.isdir(cfg["watch_folder"]) and cfg["base_folder"]:
        return

    dlg = tk.Toplevel(parent_app)
    dlg.title("Railway File Organiser — First Time Setup")
    dlg.configure(bg=C["bg"])
    dlg.geometry("560x400")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)

    sw = parent_app.winfo_screenwidth()
    sh = parent_app.winfo_screenheight()
    dlg.geometry(f"560x400+{(sw-560)//2}+{(sh-400)//2}")

    tk.Frame(dlg, bg=C["accent"], height=5).pack(fill="x")

    hdr = tk.Frame(dlg, bg=C["bg"], padx=24, pady=16)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Welcome to Railway File Organiser",
             bg=C["bg"], fg=C["text"],
             font=("Segoe UI", 15, "bold")).pack(anchor="w")
    tk.Label(hdr, text="Select two folders below. This is a one-time setup.",
             bg=C["bg"], fg=C["text2"],
             font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 0))

    tk.Frame(dlg, bg=C["border"], height=1).pack(fill="x", padx=20)

    watch_var = tk.StringVar(value=str(Path.home() / "Downloads"))
    base_var  = tk.StringVar(value=str(Path.home() / "Documents" / "Railway Files"))

    def folder_row(label, hint, var, cmd):
        f = tk.Frame(dlg, bg=C["bg"], padx=24, pady=8)
        f.pack(fill="x")
        tk.Label(f, text=label, bg=C["bg"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(f, text=hint, bg=C["bg"], fg=C["text3"],
                 font=FONT_SMALL).pack(anchor="w")
        r = tk.Frame(f, bg=C["bg"])
        r.pack(fill="x", pady=(4, 0))
        tk.Entry(r, textvariable=var, bg=C["surface2"], fg=C["text"],
                 relief="flat", font=("Segoe UI", 9),
                 insertbackground=C["text"]).pack(side="left", fill="x",
                                                   expand=True, ipady=7, padx=(0, 8))
        tk.Button(r, text="Browse", bg=C["border"], fg=C["text"],
                  font=FONT_SMALL, relief="flat", padx=10, pady=5,
                  cursor="hand2", command=cmd).pack(side="left")

    folder_row("Watch Folder — where downloads appear",
               "Usually your Downloads folder",
               watch_var,
               lambda: watch_var.set(
                   filedialog.askdirectory(title="Select Watch Folder",
                                           initialdir=watch_var.get(),
                                           parent=dlg) or watch_var.get()))

    folder_row("Organised Files Folder — where sorted files go",
               "e.g. Documents\\Railway Files",
               base_var,
               lambda: base_var.set(
                   filedialog.askdirectory(title="Select Organised Files Folder",
                                           initialdir=str(Path.home()),
                                           parent=dlg) or base_var.get()))

    def finish():
        w = watch_var.get().strip()
        b = base_var.get().strip()
        if not w or not b:
            messagebox.showwarning("Required", "Please select both folders.", parent=dlg)
            return
        cfg["watch_folder"] = w
        cfg["base_folder"]  = b
        cfg.save()
        Path(b).mkdir(parents=True, exist_ok=True)
        dlg.destroy()
        parent_app.after(300, parent_app._auto_start_if_configured)

    btn_frame = tk.Frame(dlg, bg=C["bg"], padx=24, pady=20)
    btn_frame.pack(fill="x")
    tk.Button(btn_frame, text="  Get Started  \u27a4",
              bg=C["green"], fg="#FFFFFF",
              font=("Segoe UI", 12, "bold"),
              relief="flat", padx=24, pady=10,
              cursor="hand2", command=finish).pack(side="left")
    tk.Label(btn_frame, text="You can change these later in the app.",
             bg=C["bg"], fg=C["text3"],
             font=FONT_SMALL).pack(side="left", padx=16)

    dlg.protocol("WM_DELETE_WINDOW", lambda: None)
    dlg.grab_set()
    # Do NOT call dlg.mainloop() — parent mainloop is already running!


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def main():
    # ONLY ONE tk.Tk() — all dialogs must be tk.Toplevel on this root
    app = RailwayApp()
    cfg = app.cfg
    if (not os.path.isdir(cfg["watch_folder"])) or (not cfg["base_folder"]):
        app.after(200, lambda: first_run(app, cfg))
    app.mainloop()


if __name__ == "__main__":
    main()