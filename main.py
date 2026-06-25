"""
main.py
-------
Railway File Optimiser — Main entry point.
Runs as a Windows system tray application.
Orchestrates all components: watcher → categoriser → dialog → mover → logger.
100% offline — no internet required.
"""

import os
import sys
import json
import logging
import threading
import queue
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime

# ── Internal modules ──────────────────────────────────────────────────────────
from categoriser import Categoriser
from file_watcher import FileWatcher
from file_mover import FileMover
from excel_logger import ExcelLogger
from verify_dialog import VerifyDialog

# ── Optional: system tray ─────────────────────────────────────────────────────
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

# ── Logging setup ─────────────────────────────────────────────────────────────
LOG_FILE = Path(__file__).parent / "railway_optimiser.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("Main")

# ── Config paths ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
CATEGORIES_PATH = BASE_DIR / "categories.json"

# ── Theme colours ─────────────────────────────────────────────────────────────
BG_DARK     = "#0F172A"
BG_CARD     = "#1E293B"
ACCENT      = "#3B82F6"
ACCENT_GRN  = "#22C55E"
TEXT_PRIMARY= "#F1F5F9"
TEXT_MUTED  = "#94A3B8"
BORDER      = "#475569"


# ═══════════════════════════════════════════════════════════════════════════════
class Config:
    """Load and save user configuration."""

    DEFAULTS = {
        "watch_folder": str(Path.home() / "Downloads"),
        "base_folder": str(Path.home() / "Documents" / "Railway Files"),
        "excel_log_filename": "Railway_Files_Log.xlsx",
        "auto_confirm_timeout_seconds": 60,
        "auto_start_with_windows": False,
        "open_excel_after_update": True,
        "handle_duplicates": "rename",
        "show_notifications": True,
        "min_file_size_bytes": 100,
        "stable_wait_seconds": 2,
        "theme": "dark"
    }

    def __init__(self):
        self._data = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                self._data.update(loaded)
                # Fill in watch/base folders if blank
                if not self._data["watch_folder"]:
                    self._data["watch_folder"] = self.DEFAULTS["watch_folder"]
                if not self._data["base_folder"]:
                    self._data["base_folder"] = self.DEFAULTS["base_folder"]
            except Exception as e:
                logger.warning(f"Could not load config: {e}. Using defaults.")

    def save(self):
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=4)
        except Exception as e:
            logger.error(f"Could not save config: {e}")

    def __getitem__(self, key):
        return self._data.get(key, self.DEFAULTS.get(key))

    def __setitem__(self, key, value):
        self._data[key] = value

    def get_log_path(self) -> str:
        return str(Path(self._data["base_folder"]) / self._data["excel_log_filename"])


# ═══════════════════════════════════════════════════════════════════════════════
class RailwayOptimiser:
    """Main application controller."""

    def __init__(self):
        self.config = Config()
        self.categoriser = Categoriser(CATEGORIES_PATH)
        self.watcher = None
        self.is_watching = False
        self._file_queue = queue.Queue()   # UI-thread-safe queue for dialogs
        self._setup_components()

    def _setup_components(self):
        self.mover = FileMover(self.config["base_folder"])
        log_path = self.config.get_log_path()
        self.excel = ExcelLogger(log_path,
                                  open_after_update=self.config["open_excel_after_update"])

    def start_watching(self):
        """Start the file watcher."""
        watch_folder = self.config["watch_folder"]
        if not os.path.isdir(watch_folder):
            logger.error(f"Watch folder does not exist: {watch_folder}")
            return False

        self.watcher = FileWatcher(
            watch_folder,
            on_file_ready=self._on_file_detected,
            stable_wait=float(self.config["stable_wait_seconds"]),
            stable_count=2
        )
        self.watcher.start()
        self.is_watching = True
        logger.info(f"Started watching: {watch_folder}")
        return True

    def stop_watching(self):
        """Stop the file watcher."""
        if self.watcher:
            self.watcher.stop()
            self.watcher = None
        self.is_watching = False
        logger.info("Stopped watching.")

    def _on_file_detected(self, filepath: str):
        """
        Called by FileWatcher when a new file is stable and ready.
        Runs in background thread — puts work into queue for main thread.
        """
        logger.info(f"Processing: {filepath}")
        self._process_file(filepath)

    def _process_file(self, filepath: str):
        """Full pipeline: categorise → verify → move → log."""
        filename = Path(filepath).name

        # 1. Categorise
        result = self.categoriser.categorise(filepath)
        suggested = result["category"]
        confidence = result["confidence"]

        # 2. User verification (must run on main thread if GUI)
        all_cats = self.categoriser.get_all_categories()
        dialog = VerifyDialog(
            filename=filename,
            suggested_category=suggested,
            all_categories=all_cats,
            confidence=confidence,
            timeout=self.config["auto_confirm_timeout_seconds"]
        )
        user_result = dialog.show()

        if user_result["action"] == "skip":
            logger.info(f"User skipped: {filename}")
            return

        chosen_category = user_result["category"]

        # 3. Move file
        move_result = self.mover.move(filepath, chosen_category)
        if not move_result["success"]:
            logger.error(f"Move failed: {move_result['error']}")
            self._show_error(f"Could not move file:\n{move_result['error']}")
            return

        logger.info(f"Moved to: {move_result['destination']}")

        # 4. Update Excel log
        log_result = self.excel.log(
            original_filename=move_result["original_filename"],
            final_filename=move_result["final_filename"],
            category=chosen_category,
            file_destination=move_result["destination"]
        )

        if not log_result["success"]:
            logger.warning(f"Excel log failed: {log_result['error']}")
            self._show_error(f"File moved successfully but Excel update failed:\n{log_result['error']}")
        else:
            logger.info(f"Excel updated: {log_result['log_path']}")

    def _show_error(self, message: str):
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Railway File Organiser — Error", message)
            root.destroy()
        except Exception:
            logger.error(message)

    def reload_config(self):
        """Reload config and reinitialise components."""
        self.config = Config()
        self._setup_components()

    def open_excel_log(self):
        """Open the Excel log file."""
        self.excel.open_log()

    def open_watch_folder(self):
        """Open the watch folder in Explorer."""
        try:
            os.startfile(self.config["watch_folder"])
        except Exception as e:
            logger.error(e)

    def open_base_folder(self):
        """Open the organised files base folder."""
        try:
            base = self.config["base_folder"]
            Path(base).mkdir(parents=True, exist_ok=True)
            os.startfile(base)
        except Exception as e:
            logger.error(e)


# ═══════════════════════════════════════════════════════════════════════════════
class SettingsWindow:
    """Settings GUI window."""

    def __init__(self, app: RailwayOptimiser, on_close=None):
        self.app = app
        self.on_close = on_close
        self._build()

    def _build(self):
        self.root = tk.Tk()
        self.root.title("⚙ Railway File Organiser — Settings")
        self.root.configure(bg=BG_DARK)
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # Centre
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - 600) // 2
        y = (sh - 500) // 2
        self.root.geometry(f"600x500+{x}+{y}")

        # Header
        tk.Frame(self.root, bg=ACCENT, height=6).pack(fill="x")

        title = tk.Frame(self.root, bg=BG_DARK, pady=16)
        title.pack(fill="x", padx=24)
        tk.Label(title, text="⚙  Settings", bg=BG_DARK, fg=TEXT_PRIMARY,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")

        # ── Form ──────────────────────────────────────────────────────────────
        form = tk.Frame(self.root, bg=BG_DARK)
        form.pack(fill="both", expand=True, padx=24)

        def label(text, row):
            tk.Label(form, text=text, bg=BG_DARK, fg=TEXT_MUTED,
                     font=("Segoe UI", 9, "bold")).grid(
                         row=row, column=0, sticky="w", pady=(10, 2))

        def entry_row(var, row, browse_cmd=None):
            f = tk.Frame(form, bg=BG_DARK)
            f.grid(row=row, column=0, sticky="ew", pady=(0, 4))
            form.columnconfigure(0, weight=1)
            e = tk.Entry(f, textvariable=var, bg=BG_CARD, fg=TEXT_PRIMARY,
                         insertbackground=TEXT_PRIMARY, relief="flat",
                         font=("Segoe UI", 10), width=48)
            e.pack(side="left", ipady=6, padx=(0, 8))
            if browse_cmd:
                tk.Button(f, text="Browse…", bg=ACCENT, fg=TEXT_PRIMARY,
                          font=("Segoe UI", 9), relief="flat", padx=8,
                          command=browse_cmd).pack(side="left")

        # Watch folder
        self.watch_var = tk.StringVar(value=self.app.config["watch_folder"])
        label("Watch Folder (new downloads appear here):", 0)
        entry_row(self.watch_var, 1, lambda: self._browse(self.watch_var))

        # Base folder
        self.base_var = tk.StringVar(value=self.app.config["base_folder"])
        label("Base Folder (organised files go here):", 2)
        entry_row(self.base_var, 3, lambda: self._browse(self.base_var))

        # Excel filename
        self.excel_var = tk.StringVar(value=self.app.config["excel_log_filename"])
        label("Excel Log Filename:", 4)
        entry_row(self.excel_var, 5)

        # Timeout
        self.timeout_var = tk.IntVar(value=self.app.config["auto_confirm_timeout_seconds"])
        label("Auto-confirm timeout (seconds):", 6)
        tk.Spinbox(form, textvariable=self.timeout_var, from_=10, to=300, width=6,
                   bg=BG_CARD, fg=TEXT_PRIMARY, relief="flat",
                   font=("Segoe UI", 10)).grid(row=7, column=0, sticky="w")

        # Checkboxes
        self.open_excel_var = tk.BooleanVar(value=self.app.config["open_excel_after_update"])
        tk.Checkbutton(form, text="Open Excel automatically after each file is organised",
                       variable=self.open_excel_var, bg=BG_DARK, fg=TEXT_PRIMARY,
                       selectcolor=BG_CARD, activebackground=BG_DARK,
                       font=("Segoe UI", 10)).grid(row=8, column=0, sticky="w", pady=(16, 0))

        # ── Save button ───────────────────────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=BG_DARK, pady=16)
        btn_frame.pack(fill="x", padx=24)
        tk.Button(btn_frame, text="💾  Save Settings",
                  bg=ACCENT_GRN, fg="#0F172A",
                  font=("Segoe UI", 11, "bold"), relief="flat",
                  padx=18, pady=8, command=self._save).pack(side="left")

        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.mainloop()

    def _browse(self, var: tk.StringVar):
        chosen = filedialog.askdirectory(title="Select Folder",
                                          initialdir=var.get() or str(Path.home()))
        if chosen:
            var.set(chosen)

    def _save(self):
        self.app.config["watch_folder"] = self.watch_var.get().strip()
        self.app.config["base_folder"] = self.base_var.get().strip()
        self.app.config["excel_log_filename"] = self.excel_var.get().strip()
        self.app.config["auto_confirm_timeout_seconds"] = self.timeout_var.get()
        self.app.config["open_excel_after_update"] = self.open_excel_var.get()
        self.app.config.save()
        self.app.reload_config()
        messagebox.showinfo("Settings Saved", "Settings saved successfully!\nRestart watching for changes to take effect.")
        self._close()

    def _close(self):
        self.root.destroy()
        if self.on_close:
            self.on_close()


# ═══════════════════════════════════════════════════════════════════════════════
def _make_tray_icon():
    """Create a simple railway-themed icon for the system tray."""
    img = Image.new("RGBA", (64, 64), (15, 23, 42, 255))
    draw = ImageDraw.Draw(img)
    # Train symbol: rectangle body + wheels
    draw.rounded_rectangle([8, 20, 56, 48], radius=6, fill=(59, 130, 246, 255))
    draw.rectangle([12, 24, 52, 36], fill=(15, 23, 42, 200))
    draw.ellipse([10, 46, 22, 58], fill=(30, 41, 59, 255), outline=(148, 163, 184, 255), width=2)
    draw.ellipse([42, 46, 54, 58], fill=(30, 41, 59, 255), outline=(148, 163, 184, 255), width=2)
    draw.rectangle([4, 48, 60, 50], fill=(148, 163, 184, 255))
    return img


def run_with_tray(app: RailwayOptimiser):
    """Run the application with a system tray icon."""
    icon_img = _make_tray_icon()

    def toggle_watch(icon, item):
        if app.is_watching:
            app.stop_watching()
            icon.notify("Railway File Organiser", "⏸ Stopped watching for new files.")
        else:
            if app.start_watching():
                icon.notify("Railway File Organiser", f"▶ Now watching: {app.config['watch_folder']}")

    def open_settings(icon, item):
        def _run():
            SettingsWindow(app)
        threading.Thread(target=_run, daemon=True).start()

    def open_excel(icon, item):
        app.open_excel_log()

    def open_base(icon, item):
        app.open_base_folder()

    def quit_app(icon, item):
        app.stop_watching()
        icon.stop()

    menu = pystray.Menu(
        pystray.MenuItem("▶ Start / ⏸ Stop Watching", toggle_watch, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("📊 Open Excel Log", open_excel),
        pystray.MenuItem("📁 Open Organised Files Folder", open_base),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("⚙ Settings", open_settings),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("❌ Quit", quit_app)
    )

    icon = pystray.Icon("RailwayOptimiser", icon_img,
                         "Railway File Organiser", menu)

    # Auto-start watching
    threading.Thread(target=lambda: (time.sleep(1), app.start_watching()),
                     daemon=True).start()

    icon.run()


def run_headless(app: RailwayOptimiser):
    """Fallback: run without tray (just console)."""
    print("\n" + "="*60)
    print("  🚂 Railway File Organiser — Running")
    print("="*60)
    print(f"  Watch folder : {app.config['watch_folder']}")
    print(f"  Base folder  : {app.config['base_folder']}")
    print(f"  Excel log    : {app.config.get_log_path()}")
    print("  Press Ctrl+C to stop.")
    print("="*60 + "\n")

    if not app.start_watching():
        print("ERROR: Could not start watcher. Check your watch folder in config.json.")
        sys.exit(1)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        app.stop_watching()
        print("\n👋 Railway File Organiser stopped.")


# ═══════════════════════════════════════════════════════════════════════════════
def first_run_setup():
    """Show setup wizard if config has no base_folder set."""
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "🚂 Railway File Organiser — First Run Setup",
        "Welcome!\n\nPlease select two folders:\n"
        "1. Your DOWNLOADS folder (files to watch)\n"
        "2. Your ORGANISED FILES folder (where files will go)"
    )

    watch = filedialog.askdirectory(title="Select Downloads / Watch Folder",
                                     initialdir=str(Path.home() / "Downloads"))
    if not watch:
        watch = str(Path.home() / "Downloads")

    base = filedialog.askdirectory(title="Select Organised Files Folder",
                                    initialdir=str(Path.home() / "Documents"))
    if not base:
        base = str(Path.home() / "Documents" / "Railway Files")

    root.destroy()
    return watch, base


# ═══════════════════════════════════════════════════════════════════════════════
def main():
    app = RailwayOptimiser()

    # First-run setup if folders not configured
    cfg_watch = app.config["watch_folder"]
    cfg_base = app.config["base_folder"]
    if (not os.path.isdir(cfg_watch)) or (not cfg_base):
        watch, base = first_run_setup()
        app.config["watch_folder"] = watch
        app.config["base_folder"] = base
        app.config.save()
        app.reload_config()

    # Ensure base folder exists
    Path(app.config["base_folder"]).mkdir(parents=True, exist_ok=True)

    if TRAY_AVAILABLE:
        run_with_tray(app)
    else:
        run_headless(app)


if __name__ == "__main__":
    main()
