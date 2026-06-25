"""
verify_dialog.py
----------------
User verification dialog — shows the detected file and suggested category.
Allows the user to confirm, change category, or type a new one.
Built with tkinter (built-in Python, no extra install needed).
Fully offline.
"""

import tkinter as tk
from tkinter import ttk, font
import threading


# ── Theme colours ─────────────────────────────────────────────────────────────
BG_DARK      = "#0F172A"   # Deep dark navy
BG_CARD      = "#1E293B"   # Card background
BG_INPUT     = "#334155"   # Input field background
ACCENT_BLUE  = "#3B82F6"   # Primary action
ACCENT_GREEN = "#22C55E"   # Confirm
ACCENT_RED   = "#EF4444"   # Cancel / Skip
TEXT_PRIMARY = "#F1F5F9"   # Main text
TEXT_MUTED   = "#94A3B8"   # Secondary text
TEXT_YELLOW  = "#FBBF24"   # Filename highlight
BORDER       = "#475569"   # Subtle border


class VerifyDialog:
    """
    Modal dialog to let the user confirm or override file categorisation.

    Usage:
        dialog = VerifyDialog(filename, suggested_category, all_categories, timeout=60)
        result = dialog.show()
        # result = {"action": "confirm"|"skip", "category": "..."}
    """

    def __init__(self, filename: str, suggested_category: str,
                 all_categories: list, confidence: str = "high",
                 timeout: int = 60):
        self.filename = filename
        self.suggested_category = suggested_category
        self.all_categories = sorted(all_categories)
        self.confidence = confidence
        self.timeout = timeout
        self.result = {"action": "confirm", "category": suggested_category}
        self._timer_cancelled = False

    def show(self) -> dict:
        """Display the dialog and block until user responds. Returns result dict."""
        self.root = tk.Tk()
        self.root.title("Railway File Organiser")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(False, False)

        # Centre on screen
        self.root.update_idletasks()
        w, h = 540, 400
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        # Always on top
        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()

        self._build_ui()

        # Start countdown timer
        self._remaining = self.timeout
        self._countdown()

        self.root.protocol("WM_DELETE_WINDOW", self._on_skip)
        self.root.mainloop()
        return self.result

    def _build_ui(self):
        root = self.root

        # ── Header bar ──────────────────────────────────────────────────────
        header = tk.Frame(root, bg=ACCENT_BLUE, height=6)
        header.pack(fill="x")

        # ── Title ───────────────────────────────────────────────────────────
        title_frame = tk.Frame(root, bg=BG_DARK, pady=16)
        title_frame.pack(fill="x", padx=24)

        tk.Label(title_frame, text="📂  New File Detected",
                 bg=BG_DARK, fg=TEXT_PRIMARY,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w")

        tk.Label(title_frame, text="Please confirm or change the suggested category",
                 bg=BG_DARK, fg=TEXT_MUTED,
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        # ── Separator ───────────────────────────────────────────────────────
        tk.Frame(root, bg=BORDER, height=1).pack(fill="x", padx=24)

        # ── File info card ──────────────────────────────────────────────────
        card = tk.Frame(root, bg=BG_CARD, pady=14, padx=18,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=24, pady=14)

        tk.Label(card, text="FILE NAME", bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w")

        # Truncate long filenames for display
        display_name = (self.filename[:60] + "…") if len(self.filename) > 63 else self.filename
        tk.Label(card, text=display_name, bg=BG_CARD, fg=TEXT_YELLOW,
                 font=("Segoe UI", 10, "bold"), wraplength=460,
                 justify="left").grid(row=1, column=0, sticky="w", pady=(2, 10))

        tk.Label(card, text="SUGGESTED CATEGORY", bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w")

        confidence_color = {"high": ACCENT_GREEN, "medium": TEXT_YELLOW, "low": ACCENT_RED
                            }.get(self.confidence, TEXT_MUTED)
        conf_row = tk.Frame(card, bg=BG_CARD)
        conf_row.grid(row=3, column=0, sticky="w", pady=(2, 0))

        tk.Label(conf_row, text=f"  {self.suggested_category}  ",
                 bg=ACCENT_BLUE, fg=TEXT_PRIMARY,
                 font=("Segoe UI", 11, "bold"),
                 padx=6, pady=3).pack(side="left")
        tk.Label(conf_row, text=f"  {self.confidence.upper()} CONFIDENCE",
                 bg=BG_CARD, fg=confidence_color,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=8)

        # ── Override section ─────────────────────────────────────────────────
        override_frame = tk.Frame(root, bg=BG_DARK)
        override_frame.pack(fill="x", padx=24)

        tk.Label(override_frame, text="Change Category:",
                 bg=BG_DARK, fg=TEXT_PRIMARY,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        select_row = tk.Frame(override_frame, bg=BG_DARK)
        select_row.pack(fill="x", pady=(6, 0))

        # Dropdown of existing categories
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TCombobox",
                        fieldbackground=BG_INPUT,
                        background=BG_INPUT,
                        foreground=TEXT_PRIMARY,
                        selectbackground=ACCENT_BLUE,
                        selectforeground=TEXT_PRIMARY)

        self.selected_category = tk.StringVar(value=self.suggested_category)
        self.combo = ttk.Combobox(select_row,
                                   textvariable=self.selected_category,
                                   values=self.all_categories,
                                   style="Dark.TCombobox",
                                   font=("Segoe UI", 10),
                                   width=32,
                                   state="normal")
        self.combo.pack(side="left")

        tk.Label(select_row, text="  or type a new category →",
                 bg=BG_DARK, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(side="left")

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_frame = tk.Frame(root, bg=BG_DARK, pady=18)
        btn_frame.pack(fill="x", padx=24)

        # Confirm button
        confirm_btn = tk.Button(btn_frame,
                                text="✅  Confirm & Organise",
                                bg=ACCENT_GREEN, fg="#0F172A",
                                font=("Segoe UI", 11, "bold"),
                                relief="flat", bd=0, padx=18, pady=8,
                                cursor="hand2",
                                command=self._on_confirm)
        confirm_btn.pack(side="left", padx=(0, 12))
        confirm_btn.bind("<Enter>", lambda e: confirm_btn.config(bg="#16A34A"))
        confirm_btn.bind("<Leave>", lambda e: confirm_btn.config(bg=ACCENT_GREEN))

        # Skip button
        skip_btn = tk.Button(btn_frame,
                             text="⏭  Skip This File",
                             bg=BG_CARD, fg=TEXT_MUTED,
                             font=("Segoe UI", 10),
                             relief="flat", bd=0, padx=14, pady=8,
                             cursor="hand2",
                             command=self._on_skip)
        skip_btn.pack(side="left")

        # Countdown label
        self.countdown_label = tk.Label(btn_frame,
                                         text=f"Auto-confirm in {self.timeout}s",
                                         bg=BG_DARK, fg=TEXT_MUTED,
                                         font=("Segoe UI", 9))
        self.countdown_label.pack(side="right")

    def _countdown(self):
        if self._timer_cancelled:
            return
        if self._remaining <= 0:
            self._on_confirm()
            return
        self.countdown_label.config(text=f"Auto-confirm in {self._remaining}s")
        self._remaining -= 1
        self.root.after(1000, self._countdown)

    def _on_confirm(self):
        self._timer_cancelled = True
        chosen = self.selected_category.get().strip()
        if not chosen:
            chosen = self.suggested_category
        self.result = {"action": "confirm", "category": chosen}
        self.root.destroy()

    def _on_skip(self):
        self._timer_cancelled = True
        self.result = {"action": "skip", "category": None}
        self.root.destroy()


# ---------- Quick test ----------
if __name__ == "__main__":
    categories = [
        "Tenders", "Circulars", "Reports", "Schedules",
        "Budget & Finance", "Technical Drawings", "Correspondence",
        "Policy & Rules", "Minutes of Meeting", "Contracts & Agreements",
        "Safety & Accident", "Staff & HR", "Miscellaneous"
    ]
    dialog = VerifyDialog(
        filename="TDR_2024_0056_NIT_Mumbai_East_Division_Final.pdf",
        suggested_category="Tenders",
        all_categories=categories,
        confidence="high",
        timeout=60
    )
    result = dialog.show()
    print("User result:", result)
