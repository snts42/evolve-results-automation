"""
E-volve SecureAssess Automation - GUI
Single-page layout, light mode, City & Guilds brand colours.
Built by Alex Santonastaso (snts42)
"""

import os
import re
import json
import logging
import shutil
import time
import threading
import tkinter as tk
import queue
import webbrowser
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox

from .config import (APP_VER, ENCRYPTED_CREDENTIALS_FILE, BASE_DIR,
                     get_excel_file_for_year, get_reports_base_for_year,
                     get_logs_base_for_year, load_settings, save_settings,
                     list_year_folders, list_year_excel_files)
from .secure_credentials import SecureCredentialManager

try:
    from winotify import Notification as _WinNotification
    _HAS_WINOTIFY = True
except ImportError:
    _HAS_WINOTIFY = False

try:
    import pystray
    from PIL import Image as _PILImage
    _HAS_TRAY = True
except ImportError:
    _HAS_TRAY = False
    _PILImage = None


# =============================================================================
# CONSTANTS
# =============================================================================

APP_TITLE    = "E-volve SecureAssess Automation"
# APP_VER imported from config
AUTHOR       = "Alex Santonastaso"
REPO_URL     = "https://github.com/snts42/evolve-results-automation"
ISSUES_URL   = "https://github.com/snts42/evolve-results-automation/issues"
KOFI_URL     = "https://ko-fi.com/alexsantonastaso"

# -- City & Guilds light-mode palette -----------------------------------------
# Primary brand red (City & Guilds official red)
CG_RED         = "#e30613"
CG_RED_HOVER   = "#ff1a28"
CG_RED_LIGHT   = "#FFEBEE"

# Surfaces
BG        = "#F5F5F7"   # window background (light grey)
SURFACE   = "#FFFFFF"   # card / panel / white
ELEVATED  = "#EEEEEE"   # hover / input background
BORDER    = "#DDDDDD"   # all borders

# Text
TEXT      = "#1A1A2E"   # primary (near-black)
TEXT_MID  = "#5C5C74"   # secondary
TEXT_DIM  = "#9999AA"   # placeholder / disabled

# Semantic colours
SUCCESS        = "#2E7D32"
DANGER         = "#C62828"
DANGER_BG      = "#FFEBEE"
DANGER_LIGHT   = "#FFCDD2"
AMBER          = "#E65100"

# -- Spacing ------------------------------------------------------------------
S2=2; S4=4; S6=6; S8=8; S10=10; S12=12; S16=16; S20=20; S32=32

# -- Font detection -----------------------------------------------------------
def _detect_font():
    try:
        import tkinter as _tk, tkinter.font as _tf
        r = _tk.Tk(); r.withdraw()
        fam = _tf.families(); r.destroy()
        for f in ("Segoe UI Variable Display", "Segoe UI", "Inter", "Arial"):
            if f in fam:
                return f
    except Exception:
        pass
    return "Segoe UI"

FONT = _detect_font()
MONO = "Cascadia Code"

LAST_RUN_FILE = os.path.join(BASE_DIR, "last_run.json")
ICO_PATH = os.path.join(os.path.dirname(__file__), "app.ico")
_STARTUP_FOLDER = os.path.join(
    os.environ.get("APPDATA", ""), "Microsoft", "Windows",
    "Start Menu", "Programs", "Startup")
_STARTUP_LNK = os.path.join(_STARTUP_FOLDER, "E-volve Automation.lnk")


# =============================================================================
# GUI CLASS
# =============================================================================

class EvolveGUI:
    """Single-page GUI - light mode, City & Guilds red brand colours."""

    _W = {
        "login":0.20,"refresh":0.05,
        "filter":0.10,"hashes":0.05,"scrape":0.40,"pdfs":0.15,"close":0.05,
    }

    # ------------------------------------------------------------------ init
    def __init__(self):
        # Windows: per-monitor DPI awareness + custom app ID for taskbar icon
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "evolve.secureassess.automation")
        except Exception:
            pass

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.manager        = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
        self.master_password = None
        self.automation_thread = None
        self._automation    = None
        self._stop_event    = None
        self.log_queue      = queue.Queue()
        self._read_only     = False
        self._log_handler   = None
        self._settings_win  = None
        self._help_win      = None
        self._dd_popup      = None
        self._yr_popup      = None
        self._active_popup  = None   # (popup_attr, anchor, arrow_widget)

        self._total_accounts = 0
        self._done_accounts  = 0
        self._acct_progress  = 0.0
        self._acct_pages     = 1

        self.root = ctk.CTk()
        self.root.title(f"{APP_TITLE} - Unofficial Tool")
        self.root.configure(fg_color=BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.resizable(False, False)
        self._center_window()

        self._settings = load_settings()
        self.show_browser = ctk.BooleanVar(value=self._settings.get("show_browser", False))
        self.notifications = ctk.BooleanVar(value=self._settings.get("notifications", True))
        self._sched_enabled = ctk.BooleanVar(value=self._settings.get("schedule_enabled", False))
        self._sched_time = ctk.StringVar(value=self._settings.get("schedule_time", ""))
        self._minimize_to_tray = ctk.BooleanVar(value=self._settings.get("minimize_to_tray", False))
        self._tray_icon = None
        self._scheduler_last_fired = set()
        self._update_checked = False

        # Block CTk's after(200) default-icon override
        if os.path.exists(ICO_PATH):
            try:
                self.root.iconbitmap(ICO_PATH)
            except Exception:
                pass

        # Apply high-res icon via Windows API (bypasses tkinter entirely)
        self.root.update()
        self._set_win32_icon()

    def _set_win32_icon(self, win=None):
        """Use Windows API to set correct-size icons for title bar + taskbar."""
        if not os.path.exists(ICO_PATH):
            return
        try:
            import ctypes
            u32 = ctypes.windll.user32
            ico = os.path.abspath(ICO_PATH)
            LR_LOADFROMFILE = 0x00000010

            # Query system metrics for correct icon sizes
            big_sz  = u32.GetSystemMetrics(11)   # SM_CXICON  (32/48)
            small_sz = u32.GetSystemMetrics(49)  # SM_CXSMICON (16/20)

            hicon_big = u32.LoadImageW(
                None, ico, 1, big_sz, big_sz, LR_LOADFROMFILE)
            hicon_small = u32.LoadImageW(
                None, ico, 1, small_sz, small_sz, LR_LOADFROMFILE)

            target = win or self.root
            hwnd = u32.GetParent(target.winfo_id())
            if hicon_big:
                u32.SendMessageW(hwnd, 0x0080, 1, hicon_big)    # ICON_BIG
            if hicon_small:
                u32.SendMessageW(hwnd, 0x0080, 0, hicon_small)  # ICON_SMALL
        except Exception:
            pass

    def _apply_dialog_icon(self, win):
        """Apply the app icon to a dialog window, after CTk's 200ms override."""
        if os.path.exists(ICO_PATH):
            try:
                win.after(250, lambda: win.iconbitmap(ICO_PATH) if win.winfo_exists() else None)
                win.after(300, lambda: self._set_win32_icon(win) if win.winfo_exists() else None)
            except Exception:
                pass

    # ================================================================ helpers
    def _f(self, size=14, weight="normal"):
        return ctk.CTkFont(family=FONT, size=size, weight=weight)

    def _fm(self, size=12):
        return ctk.CTkFont(family=MONO, size=size)

    def _card(self, parent, **kw):
        d = dict(fg_color=SURFACE, corner_radius=10,
                 border_width=1, border_color=BORDER)
        d.update(kw)
        return ctk.CTkFrame(parent, **d)

    def _entry(self, parent, **kw):
        d = dict(fg_color=SURFACE, border_color=BORDER, text_color=TEXT,
                 placeholder_text_color=TEXT_DIM, border_width=1,
                 corner_radius=8, height=40, font=self._f(13))
        d.update(kw)
        return ctk.CTkEntry(parent, **d)

    def _btn_primary(self, parent, text, cmd, **kw):
        d = dict(text=text, command=cmd, fg_color=CG_RED,
                 hover_color=CG_RED_HOVER, text_color=SURFACE,
                 corner_radius=8, height=40, font=self._f(13,"bold"),
                 border_width=0)
        d.update(kw)
        return ctk.CTkButton(parent, **d)

    def _btn_secondary(self, parent, text, cmd, **kw):
        d = dict(text=text, command=cmd, fg_color=SURFACE,
                 hover_color=ELEVATED, text_color=TEXT,
                 border_color=BORDER, border_width=1,
                 corner_radius=8, height=40, font=self._f(13))
        d.update(kw)
        return ctk.CTkButton(parent, **d)

    def _divider(self, parent):
        return ctk.CTkFrame(parent, height=1, fg_color=BORDER)

    # ====================================================== LOCK CARD HELPER
    def _build_lock_card(self):
        """Build the shared lock screen / reset password card layout."""
        frame = ctk.CTkFrame(self.root, fg_color=BG)
        frame.pack(fill="both", expand=True)

        outer = ctk.CTkFrame(frame, fg_color="transparent")
        outer.place(relx=0.5, rely=0.46, anchor="center")

        card = ctk.CTkFrame(outer, fg_color=SURFACE, corner_radius=16,
                            border_width=1, border_color=BORDER, width=440)
        card.pack()

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=S32, pady=S32)

        # accent stripe
        ctk.CTkFrame(inner, height=4, fg_color=CG_RED,
                     corner_radius=2).pack(fill="x", pady=(0, S20))

        return frame, inner

    # ====================================================== LOCK SCREEN
    def _show_lock_screen(self):
        self._lock, inner = self._build_lock_card()

        ctk.CTkLabel(inner, text="E-volve SecureAssess",
                     font=self._f(22,"bold"), text_color=CG_RED
                     ).pack(pady=(0, S4))
        ctk.CTkLabel(inner,
                     text=f"Automation - Unofficial Tool {APP_VER}",
                     font=self._f(12), text_color=TEXT_MID
                     ).pack(pady=(0, S12))

        is_setup = not os.path.exists(ENCRYPTED_CREDENTIALS_FILE)
        msg = ("Create a master password to encrypt your credentials."
               if is_setup else "Enter your master password to continue.")
        ctk.CTkLabel(inner, text=msg, font=self._f(12), text_color=TEXT_MID,
                     wraplength=360).pack(pady=(0, S10))

        pw1 = self._entry(inner, placeholder_text="Master password", show="*")
        pw1.pack(fill="x", pady=(0, S10))
        pw1.focus()

        pw2 = None
        strength_bar = None
        if is_setup:
            pw2 = self._entry(inner, placeholder_text="Confirm password", show="*")
            pw2.pack(fill="x", pady=(0, S6))

            str_row = ctk.CTkFrame(inner, fg_color="transparent")
            str_row.pack(fill="x", pady=(0, S10))

            strength_bar = ctk.CTkProgressBar(
                str_row, height=4, corner_radius=2,
                fg_color=ELEVATED, progress_color=DANGER)
            strength_bar.pack(side="left", fill="x", expand=True)
            strength_bar.set(0)

            strength_lbl = ctk.CTkLabel(
                str_row, text="", font=self._f(10),
                text_color=TEXT_DIM, width=44)
            strength_lbl.pack(side="right", padx=(S6, 0))

            def _update_strength(*_):
                p = pw1.get()
                score = 0
                if len(p) >= 8: score += 1
                if any(c.isupper() for c in p) and any(c.islower() for c in p): score += 1
                if any(c.isdigit() for c in p) or any(not c.isalnum() for c in p): score += 1
                labels = {0: "Weak", 1: "Weak", 2: "Fair", 3: "Strong"}
                colours = {0: DANGER, 1: DANGER, 2: AMBER, 3: SUCCESS}
                strength_bar.configure(progress_color=colours[score])
                strength_bar.set(max(score / 3, 0.05) if p else 0)
                strength_lbl.configure(text=labels[score] if p else "",
                                       text_color=colours[score])

            pw1.bind("<KeyRelease>", _update_strength)

        err = ctk.CTkLabel(inner, text="", font=self._f(11),
                           text_color=DANGER)
        err.pack(pady=(0, S10))

        def submit():
            p = pw1.get()
            if not p:
                err.configure(text="Password cannot be empty."); return
            if is_setup:
                c = pw2.get() if pw2 else ""
                if p != c:
                    err.configure(text="Passwords do not match."); return
                self.master_password = p
                self.manager.create_empty(p)
                self._unlock(first_time=True)
            else:
                try:
                    self.manager.decrypt_credentials(p)
                    self.master_password = p
                    self._unlock(first_time=False)
                except Exception:
                    err.configure(text="Incorrect password.")
                    pw1.delete(0, "end"); pw1.focus()

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x", pady=(0, S10))
        self._btn_primary(row, "Set Password" if is_setup else "Unlock",
                          submit).pack(side="left", fill="x", expand=True,
                                       padx=(0, S8))
        self._btn_secondary(row, "Quit",
                            self.root.destroy).pack(side="right", fill="x",
                                                    expand=True, padx=(S8,0))

        pw1.bind("<Return>", lambda _: submit())
        if pw2:
            pw2.bind("<Return>", lambda _: submit())

        if not is_setup:
            ctk.CTkButton(
                inner, text="Forgot password? Reset here",
                font=self._f(11), fg_color="transparent",
                hover_color=ELEVATED, text_color=TEXT_DIM,
                border_width=0, height=24,
                command=self._show_reset_password
            ).pack(pady=(S4, 0))

            ctk.CTkButton(
                inner, text="View saved results only",
                font=self._f(11), fg_color="transparent",
                hover_color=CG_RED_LIGHT, text_color=CG_RED,
                border_width=0, height=24,
                command=self._skip_to_read_only
            ).pack(pady=(S4, 0))

    # -------------------------------------------------- reset password
    def _show_reset_password(self):
        self._lock.destroy()
        self._lock, inner = self._build_lock_card()
        ctk.CTkLabel(inner, text="Reset Password",
                     font=self._f(22,"bold"), text_color=CG_RED
                     ).pack(pady=(0, S12))
        ctk.CTkLabel(inner,
                     text="This deletes the encrypted credentials file and removes "
                          "all saved E-volve SecureAssess login accounts. You will need to set a "
                          "new master password and re-add your accounts.",
                     font=self._f(12), text_color=TEXT_MID,
                     wraplength=380, justify="left").pack(pady=(0, S10))
        ctk.CTkLabel(inner,
                     text="Your Excel spreadsheets, PDF reports, and log files "
                          "will NOT be deleted.",
                     font=self._f(12), text_color=CG_RED,
                     wraplength=380, justify="left").pack(pady=(0, S10))
        ctk.CTkLabel(inner, text="This action cannot be undone.",
                     font=self._f(12,"bold"), text_color=DANGER
                     ).pack(pady=(0, S16))
        ctk.CTkLabel(inner, text="Type DELETE to confirm:",
                     font=self._f(12), text_color=TEXT_MID, anchor="w"
                     ).pack(anchor="w", pady=(0, S6))

        confirm_entry = self._entry(inner, placeholder_text="Type DELETE here")
        confirm_entry.pack(fill="x", pady=(0, S10))
        confirm_entry.focus()

        err = ctk.CTkLabel(inner, text="", font=self._f(11), text_color=DANGER)
        err.pack(pady=(0, S10))

        def do_reset():
            if confirm_entry.get().strip() != "DELETE":
                err.configure(text="Type DELETE (all caps) to confirm."); return
            try:
                os.remove(ENCRYPTED_CREDENTIALS_FILE)
            except Exception:
                pass
            self.master_password = None
            self._lock.destroy()
            self._show_lock_screen()

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")
        self._btn_primary(row, "Reset Everything", do_reset
                          ).pack(side="left", fill="x", expand=True, padx=(0, S8))
        self._btn_secondary(row, "Go Back",
                            lambda: (self._lock.destroy(), self._show_lock_screen())
                            ).pack(side="right", fill="x", expand=True, padx=(S8,0))
        confirm_entry.bind("<Return>", lambda _: do_reset())

    # ------------------------------------------------------- return to lock
    def _return_to_lock(self):
        if self.automation_thread and self.automation_thread.is_alive():
            return
        self.master_password = None
        self._read_only = False
        self.root.unbind("<Control-s>")
        self.root.unbind("<Control-l>")
        self._main.destroy()
        self._show_lock_screen()

    # ------------------------------------------------------- skip to read-only
    def _skip_to_read_only(self):
        self._read_only = True
        self.root.title(f"{APP_TITLE} - Read Only")
        if hasattr(self, '_lock') and self._lock.winfo_exists():
            self._lock.destroy()
        self._build_main_ui()
        self._log(f"{APP_TITLE}  {APP_VER}")
        self._log("Read-only mode - unlock to run automation or manage accounts.\n")

    # ------------------------------------------------------- unlock
    def _unlock(self, first_time=False):
        self._read_only = False
        self.root.title(f"{APP_TITLE} - Unofficial Tool")
        self._lock.destroy()
        self._build_main_ui()
        self._refresh_account_data()
        self._log(f"{APP_TITLE}  {APP_VER}")
        self._log("Ready.\n")
        if first_time:
            self._log("First-time setup complete. Open Settings to add your E-volve SecureAssess login.\n")
        elif not self._account_values or self._account_values == ["All Accounts"]:
            self._log("No accounts configured. Open Settings to add one.\n")

    # ============================================================= MAIN UI
    def _build_main_ui(self):
        self._main = ctk.CTkFrame(self.root, fg_color=BG)
        self._main.pack(fill="both", expand=True)

        # Red accent stripe
        ctk.CTkFrame(self._main, height=3, fg_color=CG_RED,
                     corner_radius=0).pack(fill="x")

        # Build footer first (packs at bottom)
        self._build_footer()

        # Content area
        content = ctk.CTkFrame(self._main, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=S12, pady=(S6, 0))

        self._build_controls(content)
        self._build_quick_open(content)
        self._build_activity_log(content)

        # Ctrl+H (Help) always available regardless of lock state
        self.root.bind("<Control-h>", lambda e: self._open_help())

        # Keyboard shortcuts (only in authenticated mode)
        if not self._read_only:
            self.root.bind("<Control-s>", lambda e: self._open_settings())
            self.root.bind("<Control-l>", lambda e: self._return_to_lock()
                           if not (self.automation_thread and self.automation_thread.is_alive())
                           else None)
            if not getattr(self, '_shortcuts_hinted', False):
                self._log("Shortcuts: Ctrl+S Settings, Ctrl+H Help, Ctrl+L Lock\n")
                self._shortcuts_hinted = True

        # Background update check (log-only)
        self._check_for_updates()

    # ------------------------------------------------------- controls card
    def _build_controls(self, parent):
        ctrl_card = self._card(parent)
        ctrl_card.pack(fill="x", pady=(0, S6))

        # Header row: AUTOMATION label + Settings / Help
        hdr = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        hdr.pack(fill="x", padx=S16, pady=(S8, 0))

        ctk.CTkLabel(hdr, text="AUTOMATION",
                     font=self._f(10, "bold"), text_color=CG_RED
                     ).pack(side="left")

        ctk.CTkButton(
            hdr, text="Help", width=40, height=24, corner_radius=6,
            font=self._f(11), fg_color="transparent",
            hover_color=ELEVATED, text_color=TEXT_DIM, border_width=0,
            command=self._open_help
        ).pack(side="right", padx=(S4, 0))

        self._settings_btn = ctk.CTkButton(
            hdr, text="Settings", width=60, height=24, corner_radius=6,
            font=self._f(11), fg_color="transparent",
            hover_color=ELEVATED, text_color=TEXT_DIM, border_width=0,
            command=self._open_settings
        )
        self._settings_btn.pack(side="right")

        if self._read_only:
            self._settings_btn.configure(state="disabled")
            ctk.CTkButton(
                hdr, text="Unlock", width=50, height=24, corner_radius=6,
                font=self._f(11, "bold"), fg_color="transparent",
                hover_color=CG_RED_LIGHT, text_color=CG_RED, border_width=0,
                command=self._return_to_lock
            ).pack(side="right", padx=(0, S4))
        else:
            self._lock_btn = ctk.CTkButton(
                hdr, text="Lock", width=40, height=24, corner_radius=6,
                font=self._f(11), fg_color="transparent",
                hover_color=ELEVATED, text_color=TEXT_DIM, border_width=0,
                command=self._return_to_lock
            )
            self._lock_btn.pack(side="right", padx=(S4, 0))

        # Controls row: account selector + run button
        ctrl = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        ctrl.pack(fill="x", padx=S16, pady=S10)

        ctk.CTkLabel(ctrl, text="E-volve Account",
                     font=self._f(12), text_color=TEXT_MID
                     ).pack(side="left", padx=(0, S8))

        self.account_var = ctk.StringVar(value="All Accounts")
        self._account_values = ["All Accounts"]

        # Custom dropdown trigger: frame with left text + right arrow
        self._dd_frame = ctk.CTkFrame(
            ctrl, fg_color=SURFACE, border_color=BORDER, border_width=1,
            corner_radius=8, width=130, height=36)
        self._dd_frame.pack(side="left", padx=(0, S16))
        self._dd_frame.pack_propagate(False)

        self._dd_text = ctk.CTkLabel(
            self._dd_frame, textvariable=self.account_var,
            font=self._f(12), text_color=TEXT, anchor="w")
        self._dd_text.pack(side="left", fill="x", expand=True, padx=(S8, 0))

        self._dd_arrow = ctk.CTkLabel(
            self._dd_frame, text="\u25BE", font=self._f(12),
            text_color=TEXT_DIM, width=24)
        self._dd_arrow.pack(side="right", padx=(0, S4))

        def _dd_enter(e):
            if not self._dd_popup:
                self._dd_frame.configure(fg_color=ELEVATED, border_color=TEXT_DIM)
        def _dd_leave(e):
            if not self._dd_popup:
                self._dd_frame.configure(fg_color=SURFACE, border_color=BORDER)
        for w in (self._dd_frame, self._dd_text, self._dd_arrow):
            w.bind("<Button-1>", lambda e: self._show_account_menu())
            w.bind("<Enter>", _dd_enter)
            w.bind("<Leave>", _dd_leave)

        self.cred_lbl = ctk.CTkLabel(
            ctrl,
            text="Locked" if self._read_only else "",
            font=self._f(11), text_color=TEXT_DIM)
        self.cred_lbl.pack(side="left")

        self.run_btn = self._btn_primary(
            ctrl, "Run Automation", self._run_automation,
            height=36, width=160)
        self.run_btn.pack(side="right")

        if self._read_only:
            self.run_btn.configure(state="disabled", fg_color=ELEVATED,
                                   text_color=TEXT_DIM)
            self._dd_frame.configure(fg_color=ELEVATED)
            for w in (self._dd_frame, self._dd_text, self._dd_arrow):
                w.unbind("<Button-1>")

        # Status + progress section - always packed (reserves space),
        # but starts invisible (colours match background). Revealed by
        # _show_progress_section() when automation starts.
        self._prog_frame = ctk.CTkFrame(ctrl_card, fg_color="transparent",
                                         height=50)
        self._prog_frame.pack(fill="x", padx=S16, pady=(0, S6))
        self._prog_frame.pack_propagate(False)

        self.status_lbl = ctk.CTkLabel(
            self._prog_frame, text="", font=self._f(12),
            text_color=BG, anchor="w")
        self.status_lbl.pack(fill="x", pady=(0, S4))

        self.progress_bar = ctk.CTkProgressBar(
            self._prog_frame, height=8, corner_radius=4,
            fg_color=BG, progress_color=BG)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)
        self._prog_shown = False

        if self._read_only:
            self.status_lbl.configure(
                text="Read-only mode - unlock to run automation",
                text_color=TEXT_MID)

        # Last run label (in ctrl_card, independent of progress bar)
        self._results_lbl = ctk.CTkLabel(
            ctrl_card, text="", font=self._f(11),
            text_color=TEXT_DIM, anchor="w")

        self._load_last_run()

    # ------------------------------------------------------- quick open card
    def _build_quick_open(self, parent):
        qo_card = self._card(parent)
        qo_card.pack(fill="x", pady=(0, S6))

        inner = ctk.CTkFrame(qo_card, fg_color="transparent")
        inner.pack(fill="x", padx=S16, pady=S10)

        hdr = ctk.CTkFrame(inner, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, S4))

        ctk.CTkLabel(hdr, text="QUICK OPEN",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(side="left")

        # Year selector - discover existing year folders
        cur_year = str(datetime.now().year)
        year_dirs = list_year_folders()
        if cur_year not in year_dirs:
            year_dirs.insert(0, cur_year)
        self._qo_year = ctk.StringVar(value=cur_year)
        self._qo_values = year_dirs

        self._qo_dd = ctk.CTkFrame(
            hdr, fg_color=SURFACE, border_color=BORDER, border_width=1,
            corner_radius=8, width=80, height=36)
        self._qo_dd.pack(side="right")
        self._qo_dd.pack_propagate(False)

        qo_text = ctk.CTkLabel(
            self._qo_dd, textvariable=self._qo_year,
            font=self._f(12), text_color=TEXT, anchor="w")
        qo_text.pack(side="left", fill="x", expand=True, padx=(S8, 0))

        qo_arrow = ctk.CTkLabel(
            self._qo_dd, text="\u25BE", font=self._f(12),
            text_color=TEXT_DIM, width=24)
        qo_arrow.pack(side="right", padx=(0, S4))

        def _yr_enter(e):
            if not self._yr_popup:
                self._qo_dd.configure(fg_color=ELEVATED, border_color=TEXT_DIM)
        def _yr_leave(e):
            if not self._yr_popup:
                self._qo_dd.configure(fg_color=SURFACE, border_color=BORDER)
        for w in (self._qo_dd, qo_text, qo_arrow):
            w.bind("<Button-1>", lambda e: self._show_year_menu())
            w.bind("<Enter>", _yr_enter)
            w.bind("<Leave>", _yr_leave)

        btn_row = ctk.CTkFrame(inner, fg_color="transparent")
        btn_row.pack(fill="x")

        self._excel_btn = self._btn_secondary(btn_row, "Excel",
                            self._open_current_excel,
                            height=32, width=80)
        self._excel_btn.pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "Reports",
                            lambda: self._open_current_folder("reports"),
                            height=32, width=80).pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "Logs",
                            lambda: self._open_current_folder("logs"),
                            height=32, width=80).pack(side="left")

    # ------------------------------------------------------- activity log
    def _build_activity_log(self, parent):
        log_card = self._card(parent)
        log_card.pack(fill="both", expand=True, pady=(0, S6))

        lh = ctk.CTkFrame(log_card, fg_color="transparent")
        lh.pack(fill="x", padx=S16, pady=(S8, S4))
        ctk.CTkLabel(lh, text="AUTOMATION LOGS",
                     font=self._f(10,"bold"),
                     text_color=CG_RED).pack(side="left")

        self.log_text = ctk.CTkTextbox(
            log_card, wrap="word", state="disabled",
            corner_radius=6, font=self._fm(11),
            fg_color=BG, text_color=TEXT, border_width=0)
        self.log_text.pack(fill="both", expand=True, padx=S8, pady=(0, S8))

    # ------------------------------------------------------- footer
    def _build_footer(self):
        self._divider(self._main).pack(fill="x", side="bottom")

        ft = ctk.CTkFrame(self._main, fg_color=SURFACE,
                          height=26, corner_radius=0)
        ft.pack(fill="x", side="bottom")
        ft.pack_propagate(False)

        inner = ctk.CTkFrame(ft, fg_color="transparent")
        inner.pack(fill="both", padx=S12)

        ctk.CTkLabel(inner,
                     text=f"{APP_VER}  \u2022  Made with \u2665 by {AUTHOR}",
                     font=self._f(10), text_color=TEXT_DIM
                     ).pack(side="left", pady=S2)

        ctk.CTkButton(inner, text="Support the developer",
                      font=self._f(10), fg_color="transparent",
                      hover_color=ELEVATED, text_color=AMBER,
                      border_width=0, height=20, width=130,
                      command=lambda: webbrowser.open(KOFI_URL)
                      ).pack(side="right", pady=S2)

        ctk.CTkButton(inner, text="GitHub",
                      font=self._f(10), fg_color="transparent",
                      hover_color=ELEVATED, text_color=TEXT_DIM,
                      border_width=0, height=20, width=50,
                      command=lambda: webbrowser.open(REPO_URL)
                      ).pack(side="right", pady=S2)

    # ========================================================= DIALOG HELPER
    def _create_dialog(self, title, w=440, h=420, grab=False):
        """Create a styled dialog window with accent stripe."""
        win = ctk.CTkToplevel(self.root)
        win.title(title)
        win.resizable(False, False)
        win.configure(fg_color=BG)
        win.transient(self.root)
        if grab:
            win.grab_set()
        self._center_dialog(win, w, h)
        self._apply_dialog_icon(win)
        ctk.CTkFrame(win, height=3, fg_color=CG_RED,
                     corner_radius=0).pack(fill="x")
        return win

    # ========================================================= SETTINGS DIALOG
    def _open_settings(self):
        if self._settings_win and self._settings_win.winfo_exists():
            self._settings_win.focus()
            return

        win = self._create_dialog("Settings", w=440, h=463, grab=True)
        self._settings_win = win

        inner = ctk.CTkFrame(win, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=S16, pady=S12)

        # -- GENERAL --
        ctk.CTkLabel(inner, text="GENERAL",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S4))
        ctk.CTkSwitch(
            inner, text="Show browser during automation",
            variable=self.show_browser, command=self._persist_settings,
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER
        ).pack(anchor="w", pady=(0, S4))
        ctk.CTkSwitch(
            inner, text="Desktop notifications",
            variable=self.notifications, command=self._persist_settings,
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER
        ).pack(anchor="w", pady=(0, S4))

        startup_var = ctk.BooleanVar(value=os.path.exists(_STARTUP_LNK))
        self._settings["start_with_windows"] = startup_var.get()

        def _on_startup_toggle():
            self._settings["start_with_windows"] = startup_var.get()
            self._persist_settings()
            self._toggle_startup()

        ctk.CTkSwitch(
            inner, text="Start with Windows",
            variable=startup_var, command=_on_startup_toggle,
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER
        ).pack(anchor="w", pady=(0, S4))

        if _HAS_TRAY:
            ctk.CTkSwitch(
                inner, text="Minimize to tray on close",
                variable=self._minimize_to_tray, command=self._persist_settings,
                font=self._f(12), text_color=TEXT_MID,
                progress_color=CG_RED, button_color=SURFACE,
                button_hover_color=ELEVATED, fg_color=BORDER
            ).pack(anchor="w")

        self._divider(inner).pack(fill="x", pady=(S6, S6))

        # -- SCHEDULER --
        ctk.CTkLabel(inner, text="SCHEDULER",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S4))

        sched_row = ctk.CTkFrame(inner, fg_color="transparent")
        sched_row.pack(fill="x", pady=(0, S2))

        def _sched_entry_state(enabled):
            if enabled:
                sched_entry.configure(state="normal", fg_color=SURFACE,
                                      text_color=TEXT, border_color=BORDER)
            else:
                sched_entry.configure(state="disabled", fg_color=ELEVATED,
                                      text_color=TEXT_DIM, border_color=ELEVATED)

        sched_sw = ctk.CTkSwitch(
            sched_row, text="Run daily at:",
            variable=self._sched_enabled,
            command=lambda: [_sched_entry_state(self._sched_enabled.get()),
                             _on_sched_change()],
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER)
        sched_sw.pack(side="left")

        sched_entry = ctk.CTkEntry(
            sched_row, textvariable=self._sched_time, width=68,
            font=self._f(12), fg_color=ELEVATED, border_color=ELEVATED,
            text_color=TEXT_DIM, placeholder_text="HH:MM", justify="center",
            state="disabled")
        sched_entry.pack(side="left", padx=(S8, S6))

        ctk.CTkLabel(sched_row, text="24h", font=self._f(10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, S8))

        sched_feedback = ctk.CTkLabel(
            sched_row, text="", font=self._f(10), text_color=DANGER,
            anchor="w")
        sched_feedback.pack(side="left", fill="x", expand=True)

        _sched_save_timer = [None]

        def _on_sched_change(*_args):
            if _sched_save_timer[0]:
                inner.after_cancel(_sched_save_timer[0])
                _sched_save_timer[0] = None
            if not self._sched_enabled.get():
                self._persist_settings()
                sched_feedback.configure(text="")
                return
            t = self._sched_time.get().strip()
            if t and not self._validate_time(t):
                sched_feedback.configure(text="Invalid format", text_color=DANGER)
            else:
                self._persist_settings()
                if t:
                    sched_feedback.configure(text="Saved", text_color=SUCCESS)
                    _sched_save_timer[0] = inner.after(
                        1800, lambda: sched_feedback.configure(text="")
                        if sched_feedback.winfo_exists() else None)
                else:
                    sched_feedback.configure(text="")

        self._sched_time.trace_add("write", _on_sched_change)
        _sched_entry_state(self._sched_enabled.get())
        _on_sched_change()

        self._divider(inner).pack(fill="x", pady=(S6, S6))

        # -- ADD CREDENTIAL --
        ctk.CTkLabel(inner, text="ADD CREDENTIAL",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S6))

        form = ctk.CTkFrame(inner, fg_color="transparent")
        form.pack(fill="x", pady=(0, S4))
        acc_user = self._entry(form, placeholder_text="Username", width=95)
        acc_user.pack(side="left", padx=(0, S4))
        acc_pass = self._entry(form, placeholder_text="Password", show="*", width=95)
        acc_pass.pack(side="left", padx=(0, S4))
        self._btn_primary(form, "Add", lambda: add_acct(),
                          height=40, width=66).pack(side="left", padx=(0, S4))
        status_lbl = ctk.CTkLabel(form, text="", font=self._f(10),
                                  text_color=TEXT_DIM, anchor="w")
        status_lbl.pack(side="left", fill="x", expand=True)

        # -- SAVED CREDENTIALS --
        ctk.CTkLabel(inner, text="SAVED CREDENTIALS",
                     font=self._f(10, "bold"), text_color=CG_RED
                     ).pack(anchor="w", pady=(0, S4))

        acc_scroll = ctk.CTkScrollableFrame(
            inner, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=TEXT_DIM)
        acc_scroll.pack(fill="both", expand=True)
        acc_scroll._scrollbar.grid_remove()

        def refresh_list():
            for w in acc_scroll.winfo_children():
                w.destroy()
            try:
                creds = self.manager.list_credentials(
                    master_password=self.master_password)
                if not creds:
                    acc_scroll._scrollbar.grid_remove()
                    ctk.CTkLabel(acc_scroll,
                                 text="No accounts saved yet.",
                                 font=self._f(12), text_color=TEXT_DIM
                                 ).pack(pady=S16)
                    return
                for cred in creds:
                    uname = cred.get("username", "?")
                    row = ctk.CTkFrame(acc_scroll, fg_color=SURFACE,
                                       corner_radius=8, height=44)
                    row.pack(fill="x", pady=(0, S4))
                    row.pack_propagate(False)
                    ctk.CTkLabel(row, text=uname, font=self._f(12),
                                 text_color=TEXT).pack(side="left", padx=S12)
                    rbtn = ctk.CTkButton(
                        row, text="Remove", width=72, height=28,
                        corner_radius=8, font=self._f(10),
                        fg_color=DANGER_BG, hover_color=DANGER_LIGHT,
                        text_color=DANGER, border_color=DANGER,
                        border_width=1)
                    rbtn.pack(side="right", padx=S8)
                    self._bind_settings_remove(uname, rbtn, refresh_list,
                                               status_lbl)
                if len(creds) > 2:
                    acc_scroll._scrollbar.grid()
                else:
                    acc_scroll._scrollbar.grid_remove()
            except Exception as e:
                ctk.CTkLabel(acc_scroll, text=f"Error: {e}",
                             font=self._f(12), text_color=DANGER
                             ).pack(pady=S12)

        def add_acct():
            u = acc_user.get().strip()
            p = acc_pass.get().strip()
            if not u or not p:
                status_lbl.configure(text="Both fields required",
                                     text_color=DANGER)
                return
            try:
                ok = self.manager.add_credential(
                    u, p, master_password=self.master_password)
                if ok:
                    acc_user.delete(0, "end")
                    acc_pass.delete(0, "end")
                    self._refresh_account_data()
                    refresh_list()
                    status_lbl.configure(text="Saved",
                                         text_color=SUCCESS)
                else:
                    # Check if username exists to show correct error
                    try:
                        creds = self.manager.list_credentials(
                            master_password=self.master_password)
                        exists = any(c.get("username") == u for c in creds)
                    except Exception:
                        exists = False
                    if exists:
                        status_lbl.configure(text="Already exists",
                                             text_color=DANGER)
                    else:
                        status_lbl.configure(text="Error saving",
                                             text_color=DANGER)
            except Exception as e:
                status_lbl.configure(text="Error saving",
                                     text_color=DANGER)

        refresh_list()


    def _bind_settings_remove(self, username, btn, refresh_fn, status_lbl):
        state = {"confirmed": False}

        def reset():
            if btn.winfo_exists():
                btn.configure(text="Remove")
                state["confirmed"] = False

        def on_click():
            if not state["confirmed"]:
                state["confirmed"] = True
                btn.configure(text="Confirm?")
                self.root.after(3000, reset)
            else:
                try:
                    removed = self.manager.remove_credential(
                        username, master_password=self.master_password)
                    if removed:
                        self._refresh_account_data()
                        refresh_fn()
                        status_lbl.configure(
                            text="Removed", text_color=SUCCESS)
                    else:
                        status_lbl.configure(
                            text="Not found", text_color=DANGER)
                except Exception:
                    status_lbl.configure(
                        text="Error removing", text_color=DANGER)

        btn.configure(command=on_click)

    # ========================================================= HELP DIALOG
    def _open_help(self):
        if self._help_win and self._help_win.winfo_exists():
            self._help_win.focus()
            return

        win = self._create_dialog("Help")
        self._help_win = win

        # Bottom section (pinned to bottom)
        bottom = ctk.CTkFrame(win, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=S16, pady=(0, S12))

        ctk.CTkLabel(bottom,
                     text="This is an unofficial tool and is not affiliated with, "
                          "endorsed by, or associated with City & Guilds. E-volve "
                          "and SecureAssess are trademarks of The City and Guilds "
                          "of London Institute.",
                     font=self._f(10), text_color=TEXT_DIM, anchor="w",
                     justify="left", wraplength=390
                     ).pack(anchor="w", side="bottom")

        self._divider(bottom).pack(fill="x", side="bottom", pady=(0, S8))

        btn_row = ctk.CTkFrame(bottom, fg_color="transparent")
        btn_row.pack(fill="x", side="bottom", pady=(0, S10))
        self._btn_secondary(btn_row, "View on GitHub",
                            lambda: webbrowser.open(REPO_URL),
                            height=32, width=130).pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "Report an Issue",
                            lambda: webbrowser.open(ISSUES_URL),
                            height=32, width=130).pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "Desktop Shortcut",
                            self._create_desktop_shortcut,
                            height=32, width=130).pack(side="left")

        self._divider(bottom).pack(fill="x", side="bottom", pady=(0, S10))

        # Top section (steps)
        top = ctk.CTkFrame(win, fg_color="transparent")
        top.pack(fill="both", expand=True, padx=S16, pady=S12)

        ctk.CTkLabel(top,
                     text="Automatically downloads results and PDF reports "
                          "from E-volve SecureAssess.",
                     font=self._f(11), text_color=TEXT_MID, anchor="w",
                     justify="left", wraplength=390
                     ).pack(anchor="w", pady=(0, S8))

        ctk.CTkLabel(top, text="GETTING STARTED",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S4))

        steps = (
            "1. Keep this tool in its own folder. It saves data alongside itself.",
            "2. Set a master password, then add your logins in Settings.",
            "3. Click Run Automation. Results and PDF reports are saved by year.",
        )
        for step in steps:
            ctk.CTkLabel(top, text=step, font=self._f(12),
                         text_color=TEXT_MID, anchor="w",
                         justify="left", wraplength=390
                         ).pack(anchor="w", pady=(0, S4))

        ctk.CTkLabel(top, text="TROUBLESHOOTING",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(S6, S4))

        tips = (
            "Close Excel before running and do not modify the spreadsheet structure. Open or altered files will cause errors.",
            "Most errors are caused by E-volve loading slowly or an unstable connection. Enable 'Show browser' in Settings and check logs for details.",
        )
        for tip in tips:
            ctk.CTkLabel(top, text=f"-  {tip}", font=self._f(12),
                         text_color=TEXT_MID, anchor="w",
                         justify="left", wraplength=390
                         ).pack(anchor="w", pady=(0, S10))

    # ========================================================= SETTINGS PERSISTENCE
    def _persist_settings(self):
        """Save current GUI settings to disk."""
        self._settings["show_browser"] = self.show_browser.get()
        self._settings["notifications"] = self.notifications.get()
        self._settings["schedule_enabled"] = self._sched_enabled.get()
        self._settings["schedule_time"] = self._sched_time.get()
        self._settings["minimize_to_tray"] = self._minimize_to_tray.get()
        save_settings(self._settings)

    # ========================================================= SYSTEM TRAY
    def _init_tray(self):
        """Create (but don't start) the system tray icon."""
        if not _HAS_TRAY or not os.path.exists(ICO_PATH):
            return
        try:
            try:
                icon_img = _PILImage.open(ICO_PATH).convert("RGBA").resize(
                    (32, 32), _PILImage.LANCZOS)
            except Exception:
                icon_img = _PILImage.new("RGBA", (32, 32), (227, 6, 19, 255))
            menu = pystray.Menu(
                pystray.MenuItem("Open E-volve Automation", self._tray_open,
                                 default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Open Excel", self._tray_open_excel),
                pystray.MenuItem("Open Reports", self._tray_open_reports),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Run Now", self._tray_run,
                                 visible=lambda item: bool(self.master_password)),
                pystray.MenuItem("Exit", self._tray_exit))
            self._tray_icon = pystray.Icon(
                "evolve_automation", icon_img, APP_TITLE, menu)
        except Exception:
            self._tray_icon = None

    def _tray_tooltip(self):
        """Build the tray icon tooltip text."""
        if not self.master_password:
            return f"{APP_TITLE} - Locked"
        try:
            if os.path.exists(LAST_RUN_FILE):
                with open(LAST_RUN_FILE, "r") as fh:
                    d = json.load(fh)
                ts = d.get("timestamp", "")
                rows = d.get("new_rows", 0)
                pdfs = d.get("pdfs", 0)
                if ts:
                    return f"{APP_TITLE}\nLast run: {ts} | {rows} results, {pdfs} PDFs"
        except Exception:
            pass
        return APP_TITLE

    def _show_tray(self):
        """Show the tray icon in a background thread."""
        if self._tray_icon is None:
            self._init_tray()
        if self._tray_icon is not None:
            self._tray_icon.title = self._tray_tooltip()
            def _run_icon():
                try:
                    self._tray_icon.run()
                except Exception:
                    logging.exception("pystray icon thread error")
            threading.Thread(target=_run_icon, daemon=True).start()

    def _hide_tray(self):
        """Remove the tray icon."""
        if self._tray_icon is not None:
            try:
                self._tray_icon.stop()
                time.sleep(0.1)
            except Exception:
                pass
            self._tray_icon = None

    def _tray_open(self, icon=None, item=None):
        """Restore the main window from tray."""
        self.root.after(0, self._restore_from_tray)

    def _restore_from_tray(self):
        self._hide_tray()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_open_excel(self, icon=None, item=None):
        """Open current year Excel from tray menu."""
        def _do():
            year = datetime.now().year
            path = get_excel_file_for_year(year)
            if os.path.exists(path):
                self._open_path(path)
            else:
                messagebox.showinfo("Not Found",
                    f"No Excel file found for {year}.\n\nRun automation first.",
                    parent=None)
        self.root.after(0, _do)

    def _tray_open_reports(self, icon=None, item=None):
        """Open current year reports folder from tray menu."""
        def _do():
            year = datetime.now().year
            path = get_reports_base_for_year(year)
            if os.path.exists(path):
                self._open_path(path)
            else:
                messagebox.showinfo("Not Found",
                    f"No reports folder found for {year}.\n\nRun automation first.",
                    parent=None)
        self.root.after(0, _do)

    def _tray_run(self, icon=None, item=None):
        """Trigger automation from tray menu."""
        if not self.master_password or self._read_only:
            return
        self.root.after(0, self._run_automation)

    def _tray_exit(self, icon=None, item=None):
        """Full exit from tray menu."""
        def _do_exit():
            if not self._graceful_shutdown():
                return
            self._hide_tray()
            self.root.destroy()
        self.root.after(0, _do_exit)

    # ========================================================= SCHEDULER
    def _start_scheduler(self):
        """Start the background scheduler polling thread."""
        # Pre-seed fire keys for any scheduled times already past today,
        # so a restart doesn't re-trigger a run that already happened.
        today = datetime.now().strftime("%Y-%m-%d")
        now_hm = datetime.now().strftime("%H:%M")
        sched_time = self._sched_time.get().strip()
        if self._validate_time(sched_time) and now_hm > sched_time:
            self._scheduler_last_fired.add(f"{today}_{sched_time}")
        t = threading.Thread(target=self._scheduler_loop, daemon=True)
        t.start()

    def _scheduler_loop(self):
        """Poll every 30s, fire automation when scheduled time matches current HH:MM."""
        while True:
            try:
                if self._sched_enabled.get():
                    now_hm = datetime.now().strftime("%H:%M")
                    today = datetime.now().strftime("%Y-%m-%d")
                    sched_time = self._sched_time.get().strip()
                    if self._validate_time(sched_time) and sched_time == now_hm:
                        fire_key = f"{today}_{sched_time}"
                        if fire_key not in self._scheduler_last_fired:
                            if self.master_password and not (
                                    self.automation_thread and self.automation_thread.is_alive()):
                                self._scheduler_last_fired.add(fire_key)
                                logging.info(f"Scheduled run triggered at {now_hm}")
                                self.root.after(0, self._run_automation)
                            elif not self.master_password:
                                logging.warning(
                                    f"Scheduled run at {now_hm} skipped - app is locked. "
                                    "Unlock to enable scheduled automation.")
                                self.root.after(0, lambda: self._log(
                                    f"Scheduled run at {now_hm} skipped - app is locked."))
                            elif self.automation_thread and self.automation_thread.is_alive():
                                logging.info("Scheduled run skipped - automation already in progress")
                    # Clean up old fire keys (keep only today's entries)
                    self._scheduler_last_fired = {
                        k for k in self._scheduler_last_fired if k.startswith(today)}
            except Exception:
                pass
            time.sleep(30)

    @staticmethod
    def _validate_time(t):
        """Check if string is a valid HH:MM (00:00 - 23:59)."""
        if not t or len(t) != 5 or t[2] != ':':
            return False
        try:
            h, m = int(t[:2]), int(t[3:])
            return 0 <= h <= 23 and 0 <= m <= 59
        except (ValueError, IndexError):
            return False

    # ========================================================= STARTUP SHORTCUT
    def _toggle_startup(self):
        """Create or remove Windows Startup shortcut based on current setting."""
        if self._settings.get("start_with_windows"):
            try:
                os.makedirs(_STARTUP_FOLDER, exist_ok=True)
                self._create_shortcut_at(_STARTUP_FOLDER)
            except Exception:
                pass
        else:
            try:
                if os.path.exists(_STARTUP_LNK):
                    os.remove(_STARTUP_LNK)
            except OSError:
                pass

    # ========================================================= TOAST NOTIFICATIONS
    @staticmethod
    def _notify_icon_path():
        """Return a hi-res PNG path for notifications, extracting largest frame from ICO."""
        if not os.path.exists(ICO_PATH) or not _HAS_TRAY:
            return os.path.abspath(ICO_PATH) if os.path.exists(ICO_PATH) else ""
        try:
            import tempfile
            png = os.path.join(tempfile.gettempdir(), "evolve_notify_icon.png")
            ico_mtime = os.path.getmtime(ICO_PATH)
            png_mtime = os.path.getmtime(png) if os.path.exists(png) else 0
            if not os.path.exists(png) or ico_mtime > png_mtime:
                ico = _PILImage.open(ICO_PATH)
                sizes = ico.info.get("sizes", [ico.size])
                best = max(sizes, key=lambda s: s[0])
                ico.size = best
                ico.load()
                ico.convert("RGBA").save(png, "PNG")
            return png
        except Exception:
            return os.path.abspath(ICO_PATH)

    def _notify(self, title, msg):
        """Show a Windows toast notification if enabled and winotify is available."""
        if not _HAS_WINOTIFY or not self.notifications.get():
            return
        try:
            toast = _WinNotification(
                app_id="E-volve Automation",
                title=title,
                msg=msg,
                icon=self._notify_icon_path())
            toast.build().show()
        except Exception:
            pass

    # ========================================================= AUTO UPDATE CHECK
    def _check_for_updates(self):
        """Check GitHub for a newer release tag (log-only, runs in background)."""
        if self._update_checked:
            return
        self._update_checked = True
        def _worker():
            try:
                from urllib.request import urlopen, Request
                import json as _json
                url = "https://api.github.com/repos/snts42/evolve-results-automation/releases/latest"
                req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
                resp = urlopen(req, timeout=10)
                data = _json.loads(resp.read().decode())
                resp.close()
                latest = data.get("tag_name", "")
                try:
                    lv = tuple(int(x) for x in latest.lstrip("v").split("."))
                    cv = tuple(int(x) for x in APP_VER.lstrip("v").split("."))
                except (ValueError, TypeError):
                    lv = cv = ()
                if lv > cv:
                    self.root.after(0, lambda: self._log(
                        f"Update available: {latest} (current: {APP_VER})"))
            except Exception:
                pass
        threading.Thread(target=_worker, daemon=True).start()

    # ========================================================= DESKTOP SHORTCUT
    def _create_desktop_shortcut(self):
        """Create a desktop shortcut with messagebox feedback."""
        try:
            desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
            lnk = os.path.join(desktop, "E-volve Automation.lnk")
            if os.path.exists(lnk):
                messagebox.showinfo("Desktop Shortcut",
                                    "Shortcut already exists on your Desktop.")
                return
            self._create_shortcut_at(desktop)
            messagebox.showinfo("Desktop Shortcut",
                                "Shortcut created on your Desktop.")
        except Exception as e:
            messagebox.showerror("Desktop Shortcut", f"Failed: {e}")

    def _create_shortcut_at(self, folder):
        """Create a Windows shortcut (.lnk) to this app in the given folder using VBScript."""
        import sys
        import subprocess
        import tempfile
        if getattr(sys, 'frozen', False):
            target = sys.executable
        else:
            target = os.path.abspath(sys.argv[0])
        lnk = os.path.join(folder, "E-volve Automation.lnk")
        if getattr(sys, 'frozen', False):
            ico = os.path.join(sys._MEIPASS, "evolve_results_automation", "app.ico")
        else:
            ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
        vbs = (
            f'Set s=CreateObject("WScript.Shell")\n'
            f'Set sc=s.CreateShortcut("{lnk}")\n'
            f'sc.TargetPath="{target}"\n'
            f'sc.WorkingDirectory="{os.path.dirname(target)}"\n'
            f'sc.IconLocation="{ico},0"\n'
            f'sc.Save\n'
        )
        fd, tmp = tempfile.mkstemp(suffix=".vbs")
        try:
            os.write(fd, vbs.encode())
            os.close(fd)
            subprocess.run(["cscript", "//nologo", tmp], check=True,
                           capture_output=True, timeout=10)
        finally:
            try:
                os.unlink(tmp)
            except OSError:
                pass

    # ========================================================= FILE HELPERS
    def _refresh_qo_years(self):
        """Refresh the Quick Open year dropdown to include any newly created year folders."""
        try:
            cur_year = str(datetime.now().year)
            year_dirs = list_year_folders()
            if cur_year not in year_dirs:
                year_dirs.insert(0, cur_year)
            self._qo_values = year_dirs
            if self._qo_year.get() not in year_dirs:
                self._qo_year.set(cur_year)
        except Exception:
            pass

    def _open_current_excel(self):
        year = int(self._qo_year.get()) if hasattr(self, '_qo_year') else datetime.now().year
        path = get_excel_file_for_year(year)
        if os.path.exists(path):
            self._open_path(path)
        else:
            messagebox.showinfo(
                "Not Found",
                f"No Excel file found for {year}.\n\n"
                "Run automation first to generate results.")

    def _open_current_folder(self, subfolder):
        year = int(self._qo_year.get()) if hasattr(self, '_qo_year') else datetime.now().year
        lookup = {"reports": get_reports_base_for_year, "logs": get_logs_base_for_year}
        path = lookup[subfolder](year)
        if os.path.exists(path):
            self._open_path(path)
        else:
            messagebox.showinfo(
                "Not Found",
                f"No {subfolder} folder found for {year}.\n\n"
                "Run automation first to generate results.")

    def _open_path(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            self._set_status(f"Could not open: {e}", DANGER)

    # ========================================================= LAST RUN
    def _load_last_run(self):
        try:
            if os.path.exists(LAST_RUN_FILE):
                with open(LAST_RUN_FILE, "r") as fh:
                    data = json.load(fh)
                ts = data.get("timestamp", "?")
                accts = data.get("accounts", 0)
                rows = data.get("new_rows", 0)
                pdfs = data.get("pdfs", 0)
                errs = data.get("errors", 0)
                self._results_lbl.configure(
                    text=(f"Last run: {ts}  |  {accts} account(s), "
                          f"{rows} new results, {pdfs} PDF reports, {errs} error(s)"),
                    text_color=TEXT_MID)
                self._results_lbl.pack(fill="x", padx=S16, pady=(S4, S8))
        except Exception:
            pass

    def _save_last_run(self, stats):
        try:
            data = {}
            if os.path.exists(LAST_RUN_FILE):
                with open(LAST_RUN_FILE, "r") as fh:
                    data = json.load(fh)
            data.update({
                "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "accounts": stats.accounts_processed,
                "new_rows": stats.new_rows_added,
                "pdfs": stats.pdfs_downloaded,
                "errors": stats.errors_encountered,
            })
            with open(LAST_RUN_FILE, "w") as fh:
                json.dump(data, fh)
        except Exception:
            pass

    # ========================================================= DROPDOWN POPUP
    def _show_popup(self, popup_attr, anchor, values, on_pick,
                    font_size=12, item_height=32, arrow_widget=None):
        """Generic dropdown popup anchored below a trigger widget."""
        # Close any other active popup first to prevent orphaning
        if self._active_popup and self._active_popup[0] != popup_attr:
            self._close_popup(*self._active_popup)

        # Toggle: close if already open
        existing = getattr(self, popup_attr, None)
        if existing:
            try:
                if existing.winfo_exists():
                    self._close_popup(popup_attr, anchor, arrow_widget)
                    return
            except Exception:
                setattr(self, popup_attr, None)

        # Mark trigger as active
        anchor.configure(border_color=TEXT_DIM, fg_color=ELEVATED)
        if arrow_widget:
            arrow_widget.configure(text="\u25B4")

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.configure(bg=BORDER)
        popup.wm_attributes('-topmost', True)
        setattr(self, popup_attr, popup)
        self._active_popup = (popup_attr, anchor, arrow_widget)

        frame = ctk.CTkFrame(popup, fg_color=SURFACE, corner_radius=0,
                              border_width=1, border_color=BORDER)
        frame.pack(fill="both", expand=True)

        def pick(v):
            on_pick(v)
            self._close_popup(popup_attr, anchor, arrow_widget)

        for val in values:
            ctk.CTkButton(
                frame, text=val, anchor="w",
                font=self._f(font_size), fg_color="transparent",
                hover_color=ELEVATED, text_color=TEXT,
                border_width=0, height=item_height, corner_radius=6,
                command=lambda v=val: pick(v)
            ).pack(fill="x", padx=S4, pady=S2)

        popup.update_idletasks()
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 2
        w = anchor.winfo_width()
        h = frame.winfo_reqheight()
        popup.geometry(f"{w}x{h}+{x}+{y}")

        self.root.bind('<Button-1>',
                       lambda e: self._dismiss_popup(e, popup_attr, anchor,
                                                     arrow_widget))

    def _dismiss_popup(self, event, popup_attr, anchor, arrow_widget):
        """Dismiss popup if click is outside it."""
        popup = getattr(self, popup_attr, None)
        if popup and event:
            try:
                px = popup.winfo_rootx()
                py = popup.winfo_rooty()
                pw = popup.winfo_width()
                ph = popup.winfo_height()
                if px <= event.x_root <= px+pw and py <= event.y_root <= py+ph:
                    return
            except Exception:
                pass
        self._close_popup(popup_attr, anchor, arrow_widget)

    def _close_popup(self, popup_attr, anchor, arrow_widget=None):
        """Destroy popup and reset trigger styling."""
        self.root.unbind('<Button-1>')
        popup = getattr(self, popup_attr, None)
        if popup:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass
            setattr(self, popup_attr, None)
        if self._active_popup and self._active_popup[0] == popup_attr:
            self._active_popup = None
        try:
            anchor.configure(border_color=BORDER, fg_color=SURFACE)
            if arrow_widget:
                arrow_widget.configure(text="\u25BE")
        except Exception:
            pass

    # ========================================================= ACCOUNT DROPDOWN
    def _show_account_menu(self):
        self._show_popup('_dd_popup', self._dd_frame, self._account_values,
                         self._select_account, font_size=12, item_height=32,
                         arrow_widget=self._dd_arrow)

    def _select_account(self, value):
        self.account_var.set(value)

    # ========================================================= YEAR DROPDOWN
    def _show_year_menu(self):
        """Show year selector popup matching the account dropdown style."""
        self._show_popup('_yr_popup', self._qo_dd, self._qo_values,
                         lambda v: self._qo_year.set(v), font_size=11,
                         item_height=28)

    # ========================================================= ACCOUNT DATA
    def _refresh_account_data(self):
        if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            self.cred_lbl.configure(text="0 accounts")
            self._account_values = ["All Accounts"]
            self._select_account("All Accounts")
            return
        try:
            creds = self.manager.list_credentials(
                master_password=self.master_password)
            n = len(creds)
            self.cred_lbl.configure(
                text=f"{n} account{'s' if n!=1 else ''}")
            names = [c.get("username","?") for c in creds]
            self._account_values = ["All Accounts"] + names
            # Reset selection if current account was removed
            if self.account_var.get() not in self._account_values:
                self._select_account("All Accounts")
        except Exception:
            self.cred_lbl.configure(text="Error")

    # ========================================================= LOGGING
    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_status(self, text, colour=TEXT_MID):
        self.status_lbl.configure(text=text, text_color=colour)

    def _show_progress_section(self):
        if not self._prog_shown:
            self.status_lbl.configure(text_color=TEXT_MID)
            self.progress_bar.configure(fg_color=ELEVATED, progress_color=CG_RED)
            self._prog_shown = True

    def _set_progress(self, val):
        self.progress_bar.set(max(0.0, min(1.0, val)))

    # ====================================================== AUTOMATION
    def _check_excel_locks(self):
        locked = []
        for year, p in list_year_excel_files():
            # Check for Excel's hidden ~$ lock file (fast path)
            lock_file = os.path.join(os.path.dirname(p), "~$exam_results.xlsx")
            if os.path.exists(lock_file):
                locked.append(year)
                continue
            # Try read+write access - conflicts with Excel's sharing lock
            try:
                with open(p, "a+b"):
                    pass
            except (IOError, PermissionError):
                locked.append(year)
        return locked

    def _run_automation(self):
        if self.automation_thread and self.automation_thread.is_alive():
            self._set_status("Automation is already running.", DANGER)
            return
        try:
            creds = self.manager.list_credentials(
                master_password=self.master_password)
            if not creds:
                self._set_status(
                    "No accounts configured. Open Settings to add one.",
                    DANGER)
                return
        except Exception as e:
            self._set_status(f"Error: {e}", DANGER)
            return

        locked = self._check_excel_locks()
        if locked:
            self._set_status(
                f"Close Excel first: {', '.join(locked)}/exam_results.xlsx",
                DANGER)
            return

        # Disk space check (skip dialog when running headless/from tray)
        try:
            free_mb = shutil.disk_usage(BASE_DIR).free / (1024 * 1024)
            if free_mb < 100:
                window_visible = self.root.winfo_viewable()
                if window_visible:
                    if not messagebox.askyesno(
                            "Low Disk Space",
                            f"Only {free_mb:.0f} MB free. Continue anyway?"):
                        return
                else:
                    logging.warning(f"Low disk space ({free_mb:.0f} MB) - proceeding anyway (background run)")
        except OSError:
            pass

        # Backup existing Excel files before automation
        for _yr, xls in list_year_excel_files():
            try:
                shutil.copy2(xls, xls + ".bak")
                if os.path.getsize(xls + ".bak") == 0:
                    logging.warning(f"Backup of {xls} may be corrupt (0 bytes) - check disk space")
            except OSError:
                pass

        sel = self.account_var.get()
        selected = None if sel == "All Accounts" else sel
        self._total_accounts = 1 if selected else len(creds)
        self._done_accounts  = 0
        self._acct_progress  = 0.0
        self._run_start_time = time.time()

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0","end")
        self.log_text.configure(state="disabled")

        self._stop_event = threading.Event()
        self.run_btn.configure(text="Stop", fg_color=SURFACE,
                               text_color=DANGER, border_color=DANGER,
                               border_width=1, hover_color=DANGER_BG,
                               command=self._stop_automation)
        self._excel_btn.configure(state="disabled", fg_color=ELEVATED,
                                   text_color=TEXT_DIM)
        if hasattr(self, '_lock_btn') and self._lock_btn.winfo_exists():
            self._lock_btn.configure(state="disabled", text_color=BORDER)
        self._show_progress_section()
        self._set_progress(0)
        self._set_status("Starting...", CG_RED)

        headless = not self.show_browser.get()
        self.automation_thread = threading.Thread(
            target=self._worker, args=(headless, selected), daemon=True)
        self.automation_thread.start()
        self.root.after(80, self._poll_queue)

    def _worker(self, headless, selected_username):
        try:
            from .main import EvolveAutomation

            class _QH(logging.Handler):
                def __init__(self, q):
                    super().__init__(); self.q = q
                def emit(self, record):
                    self.q.put(("log", self.format(record)))

            if self._log_handler:
                logging.getLogger().removeHandler(self._log_handler)

            h = _QH(self.log_queue)
            h.setFormatter(logging.Formatter("%(asctime)s | %(message)s",
                                             "%d/%m/%Y %H:%M:%S"))
            self._log_handler = h
            logging.getLogger().addHandler(h)

            automation = EvolveAutomation(
                headless, self.master_password, selected_username,
                stop_event=self._stop_event)
            self._automation = automation
            stats = automation.run()
            self._automation = None
            self.log_queue.put(("done", stats))
        except Exception as e:
            self.log_queue.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.log_queue.get_nowait()
                if kind == "log":
                    self._log(data)
                    self._update_progress(data)
                elif kind == "done":
                    self._on_complete(data)
                elif kind == "error":
                    self._on_error(data)
        except queue.Empty:
            pass
        if self.automation_thread and self.automation_thread.is_alive():
            self.root.after(80, self._poll_queue)
        else:
            # Final drain: catch messages the thread put just before dying
            try:
                while True:
                    kind, data = self.log_queue.get_nowait()
                    if kind == "done":
                        self._on_complete(data)
                    elif kind == "error":
                        self._on_error(data)
                    elif kind == "log":
                        self._log(data)
            except queue.Empty:
                pass

    # ----------------------------------------------- progress parsing
    def _update_progress(self, msg):
        W = self._W
        if "Starting for account" in msg:
            self._acct_progress = 0.0
            self._acct_pages = 1; self._push_progress()
            self._set_status(self._stat("Logging in..."), CG_RED)
        elif "Login submitted" in msg:
            self._acct_progress = W["login"]; self._push_progress()
            self._set_status(self._stat("Loading results..."), CG_RED)
        elif "Refreshing table" in msg:
            self._acct_progress = W["login"]+W["refresh"]
            self._push_progress()
            self._set_status(self._stat("Applying filters..."), CG_RED)
        elif "Date filter updated" in msg:
            self._acct_progress = W["login"]+W["refresh"]+W["filter"]
            self._push_progress()
            self._set_status(self._stat("Loading data..."), CG_RED)
        elif "page(s) to scrape" in msg:
            m = re.search(r"Found (\d+) page", msg)
            if m:
                self._acct_pages = max(int(m.group(1)), 1)
            base = W["login"]+W["refresh"]+W["filter"]+W["hashes"]
            self._acct_progress = base; self._push_progress()
            self._set_status(
                self._stat(f"Scraping {self._acct_pages} page(s)..."),
                CG_RED)
        elif "Processing page" in msg:
            m = re.search(r"Processing page (\d+)/(\d+)", msg)
            if m:
                pg = int(m.group(1))
                base = W["login"]+W["refresh"]+W["filter"]+W["hashes"]
                self._acct_progress = base + W["scrape"]*(pg/self._acct_pages)
                self._push_progress()
                self._set_status(
                    self._stat(f"Page {pg}/{self._acct_pages}"), CG_RED)
        elif "Processing PDF for" in msg:
            base = W["login"]+W["refresh"]+W["filter"]+W["hashes"]+W["scrape"]
            self._acct_progress = min(base+W["pdfs"]*0.5, 0.95)
            self._push_progress()
            self._set_status(self._stat("Downloading PDFs..."), CG_RED)
        elif "All done for account" in msg:
            self._acct_progress = 0.95; self._push_progress()
        elif "Chrome closed" in msg:
            self._done_accounts += 1; self._acct_progress = 0.0
            self._push_progress()

    def _stat(self, detail):
        return (f"Account {self._done_accounts+1}/"
                f"{self._total_accounts}  |  {detail}")

    def _push_progress(self):
        if self._total_accounts == 0:
            return
        overall = ((self._done_accounts + self._acct_progress)
                   / self._total_accounts)
        self._set_progress(overall)

    # ----------------------------------------------- completion / error
    def _stop_automation(self):
        if not self._stop_event:
            return
        if not messagebox.askyesno(
                "Stop Automation",
                "Stop after the current operation finishes?\n"
                "All progress will be saved."):
            return
        self._stop_event.set()
        self.run_btn.configure(state="disabled", text="Stopping...",
                               fg_color=ELEVATED, text_color=TEXT_MID)

    def _reset_controls(self):
        """Re-enable run, Excel and lock buttons after automation ends."""
        self.run_btn.configure(state="normal", text="Run Automation",
                               fg_color=CG_RED, text_color=SURFACE,
                               border_width=0,
                               command=self._run_automation)
        self._excel_btn.configure(state="normal", fg_color=SURFACE,
                                   text_color=TEXT)
        if hasattr(self, '_lock_btn') and self._lock_btn.winfo_exists():
            self._lock_btn.configure(state="normal", text_color=TEXT_DIM)

    def _on_complete(self, stats):
        if not self.root.winfo_exists():
            return
        self._reset_controls()
        self._set_progress(1.0)
        self._set_status("Completed", SUCCESS)
        elapsed = int(time.time() - self._run_start_time)
        mins, secs = divmod(elapsed, 60)
        dur = f"{mins}m {secs}s" if mins else f"{secs}s"
        self._log(f"Completed in {dur}")

        # Update results summary
        result_colour = AMBER if stats.errors_encountered > 0 else SUCCESS
        self._results_lbl.configure(
            text=(f"Completed - {stats.accounts_processed} account(s), "
                  f"{stats.new_rows_added} new results, "
                  f"{stats.pdfs_downloaded} PDF reports, "
                  f"{stats.errors_encountered} error(s)"),
            text_color=result_colour)
        if not self._results_lbl.winfo_ismapped():
            self._results_lbl.pack(fill="x", padx=S16, pady=(S4, S8))

        self._save_last_run(stats)
        if self._tray_icon is not None:
            self._tray_icon.title = self._tray_tooltip()
        self._refresh_qo_years()
        skipped = getattr(stats, 'pdfs_skipped', 0)
        skipped_txt = f", {skipped} PDF(s) skipped" if skipped else ""
        self._notify("Automation Complete",
                     f"{stats.new_rows_added} new results, "
                     f"{stats.pdfs_downloaded} PDFs"
                     f"{skipped_txt}, "
                     f"{stats.errors_encountered} error(s)")
        if skipped:
            self._log(f"Warning: {skipped} PDF(s) could not be downloaded - check logs for details.")

    def _on_error(self, msg):
        if not self.root.winfo_exists():
            return
        self._reset_controls()
        self._set_progress(0)
        self._set_status("Failed", DANGER)
        self._log(f"\nERROR: {msg}\n")
        self._notify("Automation Failed", str(msg))

    # ================================================= WINDOW MANAGEMENT
    def _graceful_shutdown(self, parent=None):
        """Stop automation and quit driver if running. Returns False if user cancels."""
        if self.automation_thread and self.automation_thread.is_alive():
            if not messagebox.askyesno(
                    "Confirm", "Automation is running. Quit anyway?",
                    parent=parent):
                return False
            if self._stop_event:
                self._stop_event.set()
            if self._automation and self._automation._driver:
                try:
                    self._automation._driver.quit()
                except Exception:
                    pass
            self.automation_thread.join(timeout=3)
        return True

    def _on_close(self):
        # Not unlocked yet - just close, no tray
        if not self.master_password and not self._read_only:
            self.root.destroy()
            return

        # Minimize to tray if enabled - even during automation (runs in background)
        if self._minimize_to_tray.get() and _HAS_TRAY:
            self.root.withdraw()
            self._show_tray()
            # First-close toast (shown once) - delayed 800ms for pystray init
            # Flag stored in last_run.json so it survives settings resets
            try:
                if os.path.exists(LAST_RUN_FILE):
                    with open(LAST_RUN_FILE, "r") as _fh:
                        _d = json.load(_fh)
                else:
                    _d = {}
            except Exception:
                _d = {}
            if not _d.get("tray_first_close_shown", False):
                _d["tray_first_close_shown"] = True
                try:
                    with open(LAST_RUN_FILE, "w") as _fh:
                        json.dump(_d, _fh)
                except Exception:
                    pass
                self.root.after(800, lambda: self._notify(
                    "Still running",
                    "E-volve Automation is in the system tray. "
                    "Right-click the tray icon to exit."))
            return

        # Not going to tray - confirm quit if automation is running
        if not self._graceful_shutdown(parent=self.root):
            return

        self._hide_tray()
        self.root.destroy()

    def _center_window(self):
        """Center window at 700x560 on the current screen."""
        w, h = 700, 560
        sx = max((self.root.winfo_screenwidth()  - w) // 2, 0)
        sy = max((self.root.winfo_screenheight() - h) // 2, 0)
        self.root.geometry(f"{w}x{h}+{sx}+{sy}")

    def _center_dialog(self, dialog, w, h):
        """Center a dialog on the parent window, clamped to the same monitor."""
        self.root.update_idletasks()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        try:
            import ctypes, ctypes.wintypes, ctypes.util

            class _RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_int), ("top", ctypes.c_int),
                             ("right", ctypes.c_int), ("bottom", ctypes.c_int)]

            class _MONITORINFO(ctypes.Structure):
                _fields_ = [("cbSize", ctypes.c_uint32),
                             ("rcMonitor", _RECT), ("rcWork", _RECT),
                             ("dwFlags", ctypes.c_uint32)]

            u32 = ctypes.windll.user32
            u32.MonitorFromPoint.restype = ctypes.c_void_p
            pt = ctypes.wintypes.POINT(px + pw // 2, py + ph // 2)
            hmon = u32.MonitorFromPoint(pt, 2)   # MONITOR_DEFAULTTONEAREST
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            u32.GetMonitorInfoW(hmon, ctypes.byref(mi))
            ml, mt, mr, mb = (mi.rcWork.left, mi.rcWork.top,
                               mi.rcWork.right, mi.rcWork.bottom)
            x = max(min(x, mr - w), ml)
            y = max(min(y, mb - h), mt)
        except Exception:
            pass
        geo = f"{w}x{h}+{x}+{y}"
        dialog.geometry(geo)
        # Re-apply at 210ms - after CTkToplevel's own 200ms geometry setup
        dialog.after(210, lambda: dialog.geometry(geo)
                     if dialog.winfo_exists() else None)

    # ================================================= ENTRY POINT
    def _acquire_instance_lock(self):
        """Prevent multiple simultaneous instances using an OS-level lock file."""
        import msvcrt
        lock_path = os.path.join(BASE_DIR, "evolve.lock")
        try:
            self._instance_lock_fh = open(lock_path, "w")
            msvcrt.locking(self._instance_lock_fh.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False

    def run(self):
        """Launch the application."""
        if not self._acquire_instance_lock():
            from tkinter import messagebox as _mb
            _mb.showerror("Already Running",
                          "E-volve Automation is already open.\n\n"
                          "Check the system tray.")
            self.root.destroy()
            return
        if os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            # Already set up - open directly in read-only, unlock via header button
            self._skip_to_read_only()
        else:
            # First-time setup - must create master password
            self._show_lock_screen()
        self._start_scheduler()
        self.root.mainloop()
