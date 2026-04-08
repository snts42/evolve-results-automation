"""
E-volve SecureAssess Automation - GUI
Single-page layout, light mode, City & Guilds brand colours.
Built by Alex Santonastaso (snts42)
"""

import os
import re
import glob
import json
import time
import threading
import queue
import webbrowser
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox

from .config import APP_VER, ENCRYPTED_CREDENTIALS_FILE, BASE_DIR, get_excel_file_for_year, get_reports_base_for_year, get_logs_base_for_year
from .secure_credentials import SecureCredentialManager


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
        self.log_queue      = queue.Queue()
        self._read_only     = False
        self._log_handler   = None
        self._settings_win  = None
        self._help_win      = None
        self._dd_popup      = None

        self._total_accounts = 0
        self._done_accounts  = 0
        self._acct_progress  = 0.0
        self._acct_pages     = 1

        self.root = ctk.CTk()
        self.root.title(f"{APP_TITLE} - Unofficial Tool")
        self.root.configure(fg_color=BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.resizable(False, False)

        self.show_browser = ctk.BooleanVar(value=False)

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
        """Apply the app icon to a dialog window."""
        if os.path.exists(ICO_PATH):
            try:
                win.after(200, lambda: win.iconbitmap(ICO_PATH))
                win.after(250, lambda: self._set_win32_icon(win))
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
        if is_setup:
            pw2 = self._entry(inner, placeholder_text="Confirm password", show="*")
            pw2.pack(fill="x", pady=(0, S10))

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
        self._read_only = False
        self._main.destroy()
        self._show_lock_screen()

    # ------------------------------------------------------- skip to read-only
    def _skip_to_read_only(self):
        self._read_only = True
        self._lock.destroy()
        self._build_main_ui()
        self._log(f"{APP_TITLE}  {APP_VER}")
        self._log("Read-only mode - unlock to run automation or manage accounts.\n")

    # ------------------------------------------------------- unlock
    def _unlock(self, first_time=False):
        self._read_only = False
        self._lock.destroy()
        self._build_main_ui()
        self._refresh_account_data()
        self._log(f"{APP_TITLE}  {APP_VER}")
        self._log("Ready.\n")
        if first_time:
            self._log("First-time setup complete. Open Settings to add your E-volve SecureAssess login.\n")

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

        for w in (self._dd_frame, self._dd_text, self._dd_arrow):
            w.bind("<Button-1>", lambda e: self._show_account_menu())

        self.cred_lbl = ctk.CTkLabel(ctrl, text="",
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

        ctk.CTkLabel(inner, text="QUICK OPEN",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S4))

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

        win = self._create_dialog("Settings", grab=True)
        self._settings_win = win

        inner = ctk.CTkFrame(win, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=S16, pady=S12)

        # -- AUTOMATION SETTINGS (top) --
        ctk.CTkLabel(inner, text="AUTOMATION SETTINGS",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S4))
        ctk.CTkSwitch(
            inner, text="Show browser during automation",
            variable=self.show_browser,
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER
        ).pack(anchor="w")

        self._divider(inner).pack(fill="x", pady=(S10, S10))

        # -- ADD CREDENTIAL --
        ctk.CTkLabel(inner, text="ADD CREDENTIAL",
                     font=self._f(10, "bold"), text_color=CG_RED,
                     anchor="w").pack(anchor="w", pady=(0, S2))
        ctk.CTkLabel(inner,
                     text="E-volve SecureAssess logins - AES-256 encrypted locally.",
                     font=self._f(11), text_color=TEXT_MID,
                     anchor="w", wraplength=400).pack(anchor="w", pady=(0, S8))

        form = ctk.CTkFrame(inner, fg_color="transparent")
        form.pack(fill="x", pady=(0, S4))
        acc_user = self._entry(form, placeholder_text="Username")
        acc_user.pack(side="left", fill="x", expand=True, padx=(0, S6))
        acc_pass = self._entry(form, placeholder_text="Password", show="*")
        acc_pass.pack(side="left", fill="x", expand=True, padx=(0, S6))
        self._btn_primary(form, "Add", lambda: add_acct(),
                          height=40, width=70).pack(side="right")

        status_lbl = ctk.CTkLabel(inner, text="", font=self._f(11),
                                  text_color=TEXT_DIM)
        status_lbl.pack(anchor="w", pady=(0, S4))

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
                if len(creds) > 3:
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
                status_lbl.configure(text="Both fields required.",
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
                    status_lbl.configure(text=f"Added '{u}'.",
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
                        status_lbl.configure(text=f"'{u}' already exists.",
                                             text_color=DANGER)
                    else:
                        status_lbl.configure(text="Failed to save credential.",
                                             text_color=DANGER)
            except Exception as e:
                status_lbl.configure(text=f"Error: {e}",
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
                            text=f"Removed '{username}'.", text_color=SUCCESS)
                    else:
                        status_lbl.configure(
                            text=f"'{username}' not found.", text_color=DANGER)
                except Exception as e:
                    status_lbl.configure(
                        text=f"Error: {e}", text_color=DANGER)

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

    # ========================================================= FILE HELPERS
    def _open_current_excel(self):
        year = datetime.now().year
        path = get_excel_file_for_year(year)
        if os.path.exists(path):
            self._open_path(path)
        else:
            messagebox.showinfo(
                "Not Found",
                f"No Excel file found for {year}.\n\n"
                "Run automation first to generate results.")

    def _open_current_folder(self, subfolder):
        year = datetime.now().year
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
            data = {
                "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "accounts": stats.accounts_processed,
                "new_rows": stats.new_rows_added,
                "pdfs": stats.pdfs_downloaded,
                "errors": stats.errors_encountered,
            }
            with open(LAST_RUN_FILE, "w") as fh:
                json.dump(data, fh)
        except Exception:
            pass

    # ========================================================= ACCOUNT DROPDOWN
    def _show_account_menu(self):
        import tkinter as tk

        # Toggle: close if already open
        if self._dd_popup:
            try:
                if self._dd_popup.winfo_exists():
                    self._close_account_menu()
                    return
            except Exception:
                self._dd_popup = None

        # Mark trigger as active
        self._dd_frame.configure(border_color=TEXT_DIM, fg_color=ELEVATED)
        self._dd_arrow.configure(text="\u25B4")

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.configure(bg=BORDER)
        self._dd_popup = popup

        frame = ctk.CTkFrame(popup, fg_color=SURFACE, corner_radius=0,
                              border_width=1, border_color=BORDER)
        frame.pack(fill="both", expand=True)

        def pick(v):
            self._select_account(v)
            self._close_account_menu()

        for val in self._account_values:
            ctk.CTkButton(
                frame, text=val, anchor="w",
                font=self._f(12), fg_color="transparent",
                hover_color=ELEVATED, text_color=TEXT,
                border_width=0, height=32, corner_radius=6,
                command=lambda v=val: pick(v)
            ).pack(fill="x", padx=S4, pady=S2)

        popup.update_idletasks()
        x = self._dd_frame.winfo_rootx()
        y = self._dd_frame.winfo_rooty() + self._dd_frame.winfo_height() + 2
        w = self._dd_frame.winfo_width()
        h = frame.winfo_reqheight()
        popup.geometry(f"{w}x{h}+{x}+{y}")

        popup.after(50, popup.focus_set)
        popup.bind("<FocusOut>",
                   lambda e: self.root.after(150, self._close_account_menu))

    def _close_account_menu(self):
        popup = self._dd_popup
        if popup:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass
            self._dd_popup = None
        try:
            self._dd_frame.configure(border_color=BORDER, fg_color=SURFACE)
            self._dd_arrow.configure(text="\u25BE")
        except Exception:
            pass

    def _select_account(self, value):
        self.account_var.set(value)

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
        for p in glob.glob(os.path.join(BASE_DIR, "*", "exam_results.xlsx")):
            year = os.path.basename(os.path.dirname(p))
            if not re.match(r'^\d{4}$', year):
                continue
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

        sel = self.account_var.get()
        selected = None if sel == "All Accounts" else sel
        self._total_accounts = 1 if selected else len(creds)
        self._done_accounts  = 0
        self._acct_progress  = 0.0
        self._run_start_time = time.time()

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0","end")
        self.log_text.configure(state="disabled")

        self.run_btn.configure(state="disabled", text="Running...",
                               fg_color=ELEVATED, text_color=TEXT_MID)
        self._excel_btn.configure(state="disabled", fg_color=ELEVATED,
                                   text_color=TEXT_DIM)
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
            import logging

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
                headless, self.master_password, selected_username)
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
    def _on_complete(self, stats):
        self.run_btn.configure(state="normal", text="Run Automation",
                               fg_color=CG_RED, text_color=SURFACE)
        self._excel_btn.configure(state="normal", fg_color=SURFACE,
                                   text_color=TEXT)
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

    def _on_error(self, msg):
        self.run_btn.configure(state="normal", text="Run Automation",
                               fg_color=CG_RED, text_color=SURFACE)
        self._excel_btn.configure(state="normal", fg_color=SURFACE,
                                   text_color=TEXT)
        self._set_progress(0)
        self._set_status("Failed", DANGER)
        self._log(f"\nERROR: {msg}\n")

    # ================================================= WINDOW MANAGEMENT
    def _on_close(self):
        if self.automation_thread and self.automation_thread.is_alive():
            if not messagebox.askyesno(
                    "Confirm", "Automation is running. Quit anyway?"):
                return
            # Quit the active Chrome driver to prevent orphaned processes
            if self._automation and self._automation._driver:
                try:
                    self._automation._driver.quit()
                except Exception:
                    pass
        self.root.destroy()

    def _center_window(self):
        """Center window at 700x520 on the current screen."""
        w, h = 700, 520
        sx = max((self.root.winfo_screenwidth()  - w) // 2, 0)
        sy = max((self.root.winfo_screenheight() - h) // 2, 0)
        self.root.geometry(f"{w}x{h}+{sx}+{sy}")

    def _center_dialog(self, dialog, w, h):
        """Center a dialog relative to the parent window (multi-monitor safe)."""
        self.root.update_idletasks()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        geo = f"{w}x{h}+{x}+{y}"
        dialog.geometry(geo)
        # Re-apply after CTkToplevel finishes its own geometry setup
        dialog.after(10, lambda: dialog.geometry(geo))

    # ================================================= ENTRY POINT
    def run(self):
        """Launch the application."""
        self._show_lock_screen()
        self.root.after(50, self._center_window)
        self.root.mainloop()
