# -*- coding: utf-8 -*-
"""
batch_organiser.py
------------------
One-time batch organiser for the PROJECT work folder.
Scans the _TELEGRAM DESKTOP dump folder, categorises every file,
shows a preview, confirms with user, moves files, and logs to Excel.

Designed specifically for the CBE Railway folder structure.
Run: python batch_organiser.py
"""

import os
import sys
import json
import shutil
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime

# ── Add parent so we can import our modules ────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from categoriser import Categoriser
from excel_logger import ExcelLogger

# ── Source and destination ────────────────────────────────────────────────────
SOURCE_FOLDER = r"\\KRILLOPAUL\Users\Paulraj\Downloads\PROJECT work\_TELEGRAM DESKTOP"
BASE_FOLDER   = r"\\KRILLOPAUL\Users\Paulraj\Downloads\PROJECT work"
EXCEL_LOG     = os.path.join(BASE_FOLDER, "_ALL FOLDERS RECORD", "Railway_Files_Log.xlsx")

# ── Map your existing folder names to category names for smart matching ────────
EXISTING_FOLDER_MAP = {
    "atm":                              "ATM",
    "booths":                           "BOOTHS AND BANNERS PROMOTIONAL KIOSK",
    "banner":                           "BOOTHS AND BANNERS PROMOTIONAL KIOSK",
    "kiosk":                            "BOOTHS AND BANNERS PROMOTIONAL KIOSK",
    "cancellation":                     "Cancellation & Diversions",
    "diversion":                        "Cancellation & Diversions",
    "catering":                         "CATERING GENERAL",
    "food":                             "CATERING GENERAL",
    "cbe issue":                        "CBE ISSUES",
    "coimbatore":                       "CBE ISSUES",
    "station profile":                  "CBE STATION PROFILE",
    "completion certificate":           "Contract completion certificates",
    "cc cert":                          "Contract completion certificates",
    "contract":                         "CONTRACTS AT CBE",
    "agreement":                        "CONTRACTS AT CBE",
    "work order":                       "CONTRACTS AT CBE",
    "contractual staff":                "CONTRACTUAL STAFF & ANTECEDENTS AT CBE",
    "antecedent":                       "CONTRACTUAL STAFF & ANTECEDENTS AT CBE",
    "csr":                              "CSR",
    "corporate social":                 "CSR",
    "dy smr":                           "DY SMR C CBE FILES",
    "feasibil":                         "FEASIBILITY REPORTS & PERMISSIONS",
    "permission":                       "FEASIBILITY REPORTS & PERMISSIONS",
    "imprest":                          "IMPREST",
    "inspection report":                "INSPECTION REPORTS",
    "inspection":                       "INSPECTION REPORTS",
    "rcr inspection":                   "RCR INSPECTION",
    "rcr":                              "RCR INSPECTION",
    "officer inspection":               "Officers Inspection Reports",
    "officers inspection":              "Officers Inspection Reports",
    "remarks":                          "REMARKS FOR OFFICERS INSPECTION",
    "letter":                           "letters",
    "lr ":                              "letters",
    " to iow":                          "LETTERS TO IOW",
    "iow":                              "LETTERS TO IOW",
    "licensee":                         "LETTERS TO LICENSEES",
    " to sd":                           "LETTERS TO SD",
    "sr.dcm":                           "LETTERS TO SR.DCM",
    "dcm":                              "LETTERS TO SR.DCM",
    " to sse":                          "LETTERS TO SSE E",
    "sse":                              "LETTERS TO SSE E",
    "lost item":                        "LOST ITEMS CCO CLAIMS",
    "lost & found":                     "LOST ITEMS CCO CLAIMS",
    "cco claim":                        "LOST ITEMS CCO CLAIMS",
    "parcel":                           "Parcel",
    "proposal":                         "Proposal and Sketches",
    "sketch":                           "Proposal and Sketches",
    "station map":                      "SKETCHES AND STATION MAP",
    "layout":                           "SKETCHES AND STATION MAP",
    "romt":                             "ROMT Layout Details",
    "rpf":                              "RPF",
    "police":                           "RPF",
    "request":                          "REQUESTS",
    "indent":                           "Requirement letter or indents",
    "requirement":                      "Requirement letter or indents",
    "rms":                              "RMS",
    "staff":                            "STAFF MATTERS",
    "transfer":                         "STAFF MATTERS",
    "promotion":                        "STAFF MATTERS",
    "t&p":                              "T&P CBE STATION",
    "t&p":                              "T&P CBE STATION",
    "tools":                            "T&P CBE STATION",
    "twps":                             "TWPS",
    "oscar":                            "oscar signage",
    "signage":                          "oscar signage",
    "fine":                             "CATERING GENERAL",
    "penalty":                          "CATERING GENERAL",
    "agency":                           "CATERING GENERAL",
    "extension":                        "CONTRACTS AT CBE",
    "approval":                         "CONTRACTS AT CBE",
    "payment":                          "IMPREST",
    "receipt":                          "IMPREST",
}

IGNORE_EXTENSIONS = {".crdownload", ".part", ".tmp", ".download"}
SKIP_FILES = {"_TELEGRAM DESKTOP RECORDS ALL.xlsx", "desktop.ini", "thumbs.db"}


def categorise_by_folder(filename: str) -> str:
    """Match file name against existing folder keywords."""
    lower = filename.lower()
    for keyword, folder in EXISTING_FOLDER_MAP.items():
        if keyword.lower() in lower:
            return folder
    return "MISCELLANEOUS FILES"


def get_destination(category: str) -> str:
    """Return destination folder path, create if needed."""
    dest = os.path.join(BASE_FOLDER, category)
    os.makedirs(dest, exist_ok=True)
    return dest


def safe_copy(src: str, dest_folder: str) -> dict:
    """Copy file to destination, handle duplicates. Original is KEPT."""
    filename = os.path.basename(src)
    dest = os.path.join(dest_folder, filename)
    if os.path.exists(dest):
        stem = Path(filename).stem
        ext = Path(filename).suffix
        counter = 1
        while os.path.exists(dest):
            dest = os.path.join(dest_folder, f"{stem}_{counter}{ext}")
            counter += 1
    try:
        shutil.copy2(src, dest)   # COPY — original file stays in place
        return {"success": True, "dest": dest, "final_name": os.path.basename(dest)}
    except Exception as e:
        return {"success": False, "dest": None, "final_name": None, "error": str(e)}


# ─────────────────────────────────────────────────────────────────────────────
# GUI Preview & Confirm
# ─────────────────────────────────────────────────────────────────────────────
BG        = "#0F172A"
BG_CARD   = "#1E293B"
ACCENT    = "#3B82F6"
GRN       = "#22C55E"
RED       = "#EF4444"
YEL       = "#FBBF24"
TEXT      = "#F1F5F9"
MUTED     = "#94A3B8"
BORDER    = "#475569"


class PreviewWindow:
    """Shows all files with their proposed categories before moving."""

    def __init__(self, file_map: list):
        """
        file_map: list of dicts: {filename, src_path, category, editable_category}
        """
        self.file_map = file_map
        self.approved = False
        self.category_vars = []

    def show(self) -> bool:
        """Show preview, return True if user clicks Proceed."""
        self.root = tk.Tk()
        self.root.title("🚂 Railway File Organiser — Batch Preview")
        self.root.configure(bg=BG)
        self.root.geometry("950x620")

        # Centre
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - 950) // 2
        y = (sh - 620) // 2
        self.root.geometry(f"950x620+{x}+{y}")

        tk.Frame(self.root, bg=ACCENT, height=5).pack(fill="x")

        # Title
        hdr = tk.Frame(self.root, bg=BG, pady=12)
        hdr.pack(fill="x", padx=20)
        tk.Label(hdr, text="📂  Batch File Organiser — Preview",
                 bg=BG, fg=TEXT, font=("Segoe UI", 15, "bold")).pack(side="left")
        tk.Label(hdr, text=f"  {len(self.file_map)} files found",
                 bg=BG, fg=MUTED, font=("Segoe UI", 11)).pack(side="left", padx=10)

        # Instruction
        tk.Label(self.root,
                 text="Review the suggested categories below. You can change any category before proceeding.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(padx=20, anchor="w")

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=20, pady=8)

        # ── Scrollable table ──────────────────────────────────────────────────
        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=20)

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=BG)
        self.scroll_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mousewheel
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # ── Column headers ────────────────────────────────────────────────────
        headers = ["#", "File Name", "Suggested Category (editable)"]
        widths   = [4, 52, 38]
        for col, (h, w) in enumerate(zip(headers, widths)):
            tk.Label(self.scroll_frame, text=h, bg=BG_CARD, fg=MUTED,
                     font=("Segoe UI", 9, "bold"),
                     width=w, anchor="w", padx=6, pady=4
                     ).grid(row=0, column=col, sticky="ew", padx=1, pady=1)

        # ── Get all unique categories for dropdown ────────────────────────────
        all_cats = sorted(set(EXISTING_FOLDER_MAP.values())) + ["MISCELLANEOUS FILES"]

        # ── File rows ─────────────────────────────────────────────────────────
        for i, item in enumerate(self.file_map, start=1):
            row_bg = "#1E293B" if i % 2 == 0 else "#162032"

            tk.Label(self.scroll_frame, text=str(i), bg=row_bg, fg=MUTED,
                     font=("Segoe UI", 9), width=4, anchor="center"
                     ).grid(row=i, column=0, sticky="ew", padx=1, pady=1)

            display = item["filename"]
            if len(display) > 60:
                display = display[:57] + "…"
            tk.Label(self.scroll_frame, text=display, bg=row_bg, fg=TEXT,
                     font=("Segoe UI", 9), width=52, anchor="w", padx=6
                     ).grid(row=i, column=1, sticky="ew", padx=1, pady=1)

            cat_var = tk.StringVar(value=item["category"])
            self.category_vars.append(cat_var)
            combo = ttk.Combobox(self.scroll_frame, textvariable=cat_var,
                                  values=all_cats, font=("Segoe UI", 9),
                                  width=36, state="normal")
            combo.grid(row=i, column=2, sticky="ew", padx=1, pady=1)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_bar = tk.Frame(self.root, bg=BG, pady=14)
        btn_bar.pack(fill="x", padx=20)

        tk.Button(btn_bar, text=f"✅  Proceed — Organise {len(self.file_map)} Files",
                  bg=GRN, fg="#0F172A", font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=20, pady=9,
                  command=self._proceed).pack(side="left", padx=(0, 12))

        tk.Button(btn_bar, text="❌  Cancel",
                  bg=BG_CARD, fg=MUTED, font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=9,
                  command=self._cancel).pack(side="left")

        self.root.protocol("WM_DELETE_WINDOW", self._cancel)
        self.root.mainloop()
        return self.approved

    def _proceed(self):
        # Write user-edited categories back to file_map
        for i, item in enumerate(self.file_map):
            item["category"] = self.category_vars[i].get().strip() or item["category"]
        self.approved = True
        self.root.destroy()

    def _cancel(self):
        self.approved = False
        self.root.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Main batch run
# ─────────────────────────────────────────────────────────────────────────────
def run_batch():
    print("\n" + "="*65)
    print("  [TRAIN] Railway File Organiser -- Batch Mode")
    print("="*65)
    print(f"  Source : {SOURCE_FOLDER}")
    print(f"  Base   : {BASE_FOLDER}")
    print()

    if not os.path.isdir(SOURCE_FOLDER):
        print(f"[ERROR] Source folder not found:")
        print(f"  {SOURCE_FOLDER}")
        input("Press Enter to exit...")
        return

    # ── Collect files ─────────────────────────────────────────────────────────
    all_files = []
    for f in os.listdir(SOURCE_FOLDER):
        if f in SKIP_FILES:
            continue
        if Path(f).suffix.lower() in IGNORE_EXTENSIONS:
            continue
        full = os.path.join(SOURCE_FOLDER, f)
        if os.path.isfile(full):
            all_files.append(f)

    if not all_files:
        print("No files found to organise in source folder.")
        input("Press Enter to exit...")
        return

    print(f"  Found {len(all_files)} files to organise.")

    # ── Categorise each file ──────────────────────────────────────────────────
    file_map = []
    for filename in all_files:
        src = os.path.join(SOURCE_FOLDER, filename)
        category = categorise_by_folder(filename)
        file_map.append({
            "filename": filename,
            "src_path": src,
            "category": category
        })

    # ── Show preview GUI ──────────────────────────────────────────────────────
    preview = PreviewWindow(file_map)
    approved = preview.show()

    if not approved:
        print("\nOperation cancelled by user.")
        return

    # ── Move files & log ──────────────────────────────────────────────────────
    print(f"\n  Organising {len(file_map)} files...")
    logger = ExcelLogger(EXCEL_LOG, open_after_update=False)

    success_count = 0
    fail_count = 0
    results = []

    for item in file_map:
        dest_folder = get_destination(item["category"])
        move_result = safe_copy(item["src_path"], dest_folder)

        if move_result["success"]:
            success_count += 1
            logger.log(
                original_filename=item["filename"],
                final_filename=move_result["final_name"],
                category=item["category"],
                file_destination=move_result["dest"]
            )
            results.append(f"  OK   {item['filename'][:55]:<55} -> {item['category']}")
        else:
            fail_count += 1
            results.append(f"  FAIL {item['filename'][:55]:<55} FAILED: {move_result.get('error','')}")

    for r in results:
        print(r)

    print(f"\n{'='*65}")
    print(f"  Done! {success_count} moved, {fail_count} failed.")
    print(f"  Excel log: {EXCEL_LOG}")
    print(f"{'='*65}\n")

    # Open Excel log
    try:
        os.startfile(EXCEL_LOG)
    except Exception:
        pass

    # Summary popup
    root2 = tk.Tk()
    root2.withdraw()
    fail_msg = f"{fail_count} files failed.\n" if fail_count else ""
    messagebox.showinfo(
        "Batch Organiser Complete",
        f"{success_count} files organised successfully!\n"
        f"{fail_msg}"
        f"Excel log has been opened."
    )
    root2.destroy()


if __name__ == "__main__":
    run_batch()
