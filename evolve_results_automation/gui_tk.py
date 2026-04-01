"""
E-volve SecureAssess Results Automation - GUI
Light mode, City & Guilds brand colours, maximised window.
Built by Alex Santonastaso (snts42)
"""

import os
import re
import glob
import threading
import queue
import webbrowser
import customtkinter as ctk
from tkinter import messagebox

from .config import ENCRYPTED_CREDENTIALS_FILE, BASE_DIR
from .secure_credentials import SecureCredentialManager


# =============================================================================
# CONSTANTS
# =============================================================================

APP_TITLE    = "E-volve SecureAssess"
APP_SUBTITLE = "Results Automation"
APP_VER      = "v1.0.0"
AUTHOR       = "Alex Santonastaso"
AUTHOR_HANDLE = "snts42"
GITHUB_URL   = "https://github.com/snts42"
REPO_URL     = "https://github.com/snts42/evolve-results-automation"
ISSUES_URL   = "https://github.com/snts42/evolve-results-automation/issues"
KOFI_URL     = "https://ko-fi.com/alexsantonastaso"

# -- City & Guilds light-mode palette -----------------------------------------
# Primary brand red (City & Guilds official red)
CG_RED       = "#e30613"
CG_RED_HOVER = "#ff1a28"
CG_RED_LIGHT = "#FFEBEE"
CG_RED_MID   = "#EF5350"

BG        = "#F5F5F7"   # window background (light grey)
SURFACE   = "#FFFFFF"   # card / panel surface
ELEVATED  = "#EEEEEE"   # hover / input background
BORDER    = "#DDDDDD"   # border

TEXT      = "#1A1A2E"   # near-black
TEXT_MID  = "#5C5C74"   # secondary text
TEXT_DIM  = "#9999AA"   # placeholder / dim

SUCCESS      = "#2E7D32"
SUCCESS_BG   = "#E8F5E9"
DANGER       = "#C62828"
DANGER_BG    = "#FFEBEE"
AMBER        = "#E65100"
AMBER_BG     = "#FFF3E0"

# -- Spacing ------------------------------------------------------------------
S2=2; S4=4; S6=6; S8=8; S10=10; S12=12; S16=16; S20=20; S24=24; S32=32

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

# -- Icon generation ----------------------------------------------------------
def _generate_icon(path):
    try:
        from PIL import Image, ImageDraw
        imgs = []
        for sz in (16, 32, 48, 64):
            img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            m = max(sz // 16, 1)
            d.rounded_rectangle([m, m, sz-m-1, sz-m-1], radius=max(sz//5,2),
                                 fill=(227, 6, 19, 255), outline=(255, 26, 40, 200))
            cx, cy = sz/2 + sz*0.03, sz/2
            h, w = sz*0.44, sz*0.38
            d.polygon([(cx-w/2, cy-h/2),(cx+w/2, cy),(cx-w/2, cy+h/2)],
                      fill=(255, 255, 255, 230))
            imgs.append(img)
        imgs[0].save(path, format="ICO",
                     sizes=[(16,16),(32,32),(48,48),(64,64)],
                     append_images=imgs[1:])
        return True
    except Exception:
        return False


# =============================================================================
# GUI CLASS
# =============================================================================

class EvolveGUI:
    """Single-window GUI — light mode, City & Guilds red brand colours."""

    _W = {
        "login":0.10,"navigate":0.08,"iframe":0.04,"refresh":0.03,
        "filter":0.10,"hashes":0.05,"scrape":0.40,"pdfs":0.15,"close":0.05,
    }

    # ------------------------------------------------------------------ init
    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.manager        = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
        self.master_password = None
        self.automation_thread = None
        self.log_queue      = queue.Queue()
        self._authenticated = False
        self._active_tab    = None
        self._log_handler   = None

        self._total_accounts = 0
        self._done_accounts  = 0
        self._acct_progress  = 0.0
        self._acct_pages     = 1
        self._acct_cur_page  = 0

        self.root = ctk.CTk()
        self.root.title(f"{APP_TITLE} – {APP_SUBTITLE}")
        self.root.configure(fg_color=BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.minsize(900, 620)

        ico = os.path.join(os.path.dirname(__file__), "app.ico")
        if not os.path.exists(ico):
            _generate_icon(ico)
        if os.path.exists(ico):
            try:
                self.root.iconbitmap(ico)
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
                 hover_color=CG_RED_HOVER, text_color="#FFFFFF",
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

    def _section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=self._f(10,"bold"),
                     text_color=TEXT_DIM, anchor="w"
                     ).pack(fill="x", padx=S16, pady=(S12, S4))

    # ====================================================== LOCK SCREEN
    def _show_lock_screen(self):
        self._lock = ctk.CTkFrame(self.root, fg_color=BG)
        self._lock.pack(fill="both", expand=True)

        outer = ctk.CTkFrame(self._lock, fg_color="transparent")
        outer.place(relx=0.5, rely=0.46, anchor="center")

        card = ctk.CTkFrame(outer, fg_color=SURFACE, corner_radius=16,
                            border_width=1, border_color=BORDER, width=440)
        card.pack()

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=S32, pady=S32)

        # accent stripe
        ctk.CTkFrame(inner, height=4, fg_color=CG_RED,
                     corner_radius=2).pack(fill="x", pady=(0, S20))

        ctk.CTkLabel(inner, text=APP_TITLE,
                     font=self._f(22,"bold"), text_color=TEXT
                     ).pack(pady=(0, S4))
        ctk.CTkLabel(inner, text=APP_SUBTITLE,
                     font=self._f(13), text_color=TEXT_MID
                     ).pack(pady=(0, S4))
        ctk.CTkLabel(inner, text=f"Unofficial Tool  {APP_VER}",
                     font=self._f(10), text_color=TEXT_DIM
                     ).pack(pady=(0, S20))

        is_setup = not os.path.exists(ENCRYPTED_CREDENTIALS_FILE)
        msg = ("Create a master password to encrypt your credentials."
               if is_setup else "Enter your master password to continue.")
        ctk.CTkLabel(inner, text=msg, font=self._f(12), text_color=TEXT_MID,
                     wraplength=360).pack(pady=(0, S16))

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

    # -------------------------------------------------- reset password
    def _show_reset_password(self):
        self._lock.destroy()
        self._lock = ctk.CTkFrame(self.root, fg_color=BG)
        self._lock.pack(fill="both", expand=True)

        outer = ctk.CTkFrame(self._lock, fg_color="transparent")
        outer.place(relx=0.5, rely=0.46, anchor="center")

        card = ctk.CTkFrame(outer, fg_color=SURFACE, corner_radius=16,
                            border_width=1, border_color=BORDER, width=460)
        card.pack()

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=S32, pady=S32)

        ctk.CTkFrame(inner, height=4, fg_color=DANGER,
                     corner_radius=2).pack(fill="x", pady=(0, S20))
        ctk.CTkLabel(inner, text="Reset Password",
                     font=self._f(20,"bold"), text_color=TEXT
                     ).pack(pady=(0, S12))
        ctk.CTkLabel(inner,
                     text="This deletes the encrypted credentials file and removes "
                          "all saved E-volve login accounts. You will need to set a "
                          "new master password and re-add your accounts.",
                     font=self._f(12), text_color=TEXT_MID,
                     wraplength=380, justify="left").pack(pady=(0, S10))
        ctk.CTkLabel(inner,
                     text="Your Excel spreadsheets, PDF reports, and log files "
                          "will NOT be deleted.",
                     font=self._f(12), text_color=SUCCESS,
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
        self._btn_primary(row, "Reset Everything", do_reset,
                          fg_color=DANGER, hover_color="#B71C1C"
                          ).pack(side="left", fill="x", expand=True, padx=(0, S8))
        self._btn_secondary(row, "Go Back",
                            lambda: (self._lock.destroy(), self._show_lock_screen())
                            ).pack(side="right", fill="x", expand=True, padx=(S8,0))
        confirm_entry.bind("<Return>", lambda _: do_reset())

    # ------------------------------------------------------- unlock
    def _unlock(self, first_time=False):
        self._authenticated = True
        self._lock.destroy()
        self._build_main_ui()
        self.root.after(50, self._maximize)
        self._refresh_account_data()
        self._log(f"{APP_TITLE} {APP_SUBTITLE}  {APP_VER}")
        self._log("Ready.\n")
        if first_time:
            self._log("First-time setup complete. Head to Accounts to add your E-volve login.\n")
            self.root.after(300, lambda: self._switch_tab("accounts"))

    # ============================================================= MAIN UI
    def _build_main_ui(self):
        self._main = ctk.CTkFrame(self.root, fg_color=BG)
        self._main.pack(fill="both", expand=True)

        self._build_tab_bar()
        self._build_footer()

        # content fills between tab bar and footer
        self._content = ctk.CTkFrame(self._main, fg_color="transparent")
        self._content.pack(fill="both", expand=True, padx=S12, pady=(S4, 0))

        # toast slot (zero height when empty)
        self._toast_slot = ctk.CTkFrame(self._content, fg_color="transparent")
        self._toast_slot.pack(fill="x")

        # --- tab frames: use pack/pack_forget (avoids place sizing issues) ---
        self._tab_dashboard = ctk.CTkFrame(self._content, fg_color="transparent")
        self._tab_accounts  = ctk.CTkFrame(self._content, fg_color="transparent")
        self._tab_files     = ctk.CTkFrame(self._content, fg_color="transparent")
        self._tab_help      = ctk.CTkFrame(self._content, fg_color="transparent")

        self._build_dashboard()
        self._build_accounts()
        self._build_files()
        self._build_help()

        self._switch_tab("dashboard")

    # ------------------------------------------------------- tab bar
    def _build_tab_bar(self):
        # Red accent line
        ctk.CTkFrame(self._main, height=4, fg_color=CG_RED,
                     corner_radius=0).pack(fill="x")

        bar = ctk.CTkFrame(self._main, fg_color=SURFACE,
                           height=44, corner_radius=0)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", padx=S16, pady=S4)

        # Label on right
        ctk.CTkLabel(inner,
                     text=f"{APP_TITLE} {APP_SUBTITLE}  —  Unofficial Tool  {APP_VER}",
                     font=self._f(11), text_color=TEXT_DIM
                     ).pack(side="right")

        self._tab_btns = {}
        self._tab_indicators = {}

        for name, label in [("dashboard","Dashboard"),("accounts","Accounts"),
                             ("files","Files"),("help","Help")]:
            col = ctk.CTkFrame(inner, fg_color="transparent")
            col.pack(side="left", padx=(0, S4))

            btn = ctk.CTkButton(
                col, text=label, font=self._f(12),
                fg_color="transparent", hover_color=CG_RED_LIGHT,
                text_color=TEXT_MID, corner_radius=6,
                height=28, width=90, border_width=0,
                command=lambda n=name: self._switch_tab(n))
            btn.pack()

            ind = ctk.CTkFrame(col, height=2, fg_color="transparent",
                               corner_radius=1)
            ind.pack(fill="x")

            self._tab_btns[name] = btn
            self._tab_indicators[name] = ind

        self._divider(self._main).pack(fill="x")

    # ------------------------------------------------------- tab switch
    def _switch_tab(self, name):
        frames = {
            "dashboard": self._tab_dashboard,
            "accounts":  self._tab_accounts,
            "files":     self._tab_files,
            "help":      self._tab_help,
        }
        for f in frames.values():
            f.pack_forget()

        frames[name].pack(fill="both", expand=True)

        for n, btn in self._tab_btns.items():
            active = (n == name)
            btn.configure(text_color=CG_RED if active else TEXT_MID,
                          font=self._f(12, "bold" if active else "normal"))
            self._tab_indicators[n].configure(
                fg_color=CG_RED if active else "transparent")

        self._active_tab = name
        if name == "accounts":
            self._refresh_accounts_list()
        elif name == "files":
            self._refresh_files_list()

    # ------------------------------------------------------- footer
    def _build_footer(self):
        ft = ctk.CTkFrame(self._main, fg_color=SURFACE,
                          height=40, corner_radius=0)
        ft.pack(fill="x", side="bottom")
        ft.pack_propagate(False)

        self._divider(self._main).pack(fill="x", side="bottom")

        inner = ctk.CTkFrame(ft, fg_color="transparent")
        inner.pack(fill="both", padx=S16)

        ctk.CTkLabel(inner, text=f"Made with ♥ by {AUTHOR}",
                     font=self._f(12), text_color=TEXT_MID
                     ).pack(side="left", pady=S6)

        ctk.CTkButton(inner, text="Support on Ko-fi",
                      font=self._f(12), fg_color="transparent",
                      hover_color=ELEVATED, text_color=AMBER,
                      border_width=0, height=28, width=130,
                      command=lambda: webbrowser.open(KOFI_URL)
                      ).pack(side="right", pady=S6, padx=(S8,0))

        ctk.CTkButton(inner, text=f"github.com/{AUTHOR_HANDLE}",
                      font=self._f(12), fg_color="transparent",
                      hover_color=ELEVATED, text_color=TEXT_MID,
                      border_width=0, height=28, width=140,
                      command=lambda: webbrowser.open(GITHUB_URL)
                      ).pack(side="right", pady=S6)

    # ========================================================= DASHBOARD
    def _build_dashboard(self):
        f = self._tab_dashboard

        # --- controls row ---
        ctrl_card = self._card(f)
        ctrl_card.pack(fill="x", pady=(0, S6))

        ctrl = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        ctrl.pack(fill="x", padx=S16, pady=S10)

        ctk.CTkLabel(ctrl, text="Account",
                     font=self._f(12), text_color=TEXT_MID
                     ).pack(side="left", padx=(0, S8))

        self.account_var = ctk.StringVar(value="All Accounts")
        self.account_dd = ctk.CTkOptionMenu(
            ctrl, variable=self.account_var,
            values=["All Accounts"],
            height=36, corner_radius=8, font=self._f(12), width=200,
            fg_color=SURFACE, button_color=BORDER,
            button_hover_color=ELEVATED, text_color=TEXT,
            dropdown_fg_color=SURFACE, dropdown_hover_color=CG_RED_LIGHT,
            dropdown_text_color=TEXT)
        self.account_dd.pack(side="left", padx=(0, S16))

        self.show_browser = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            ctrl, text="Show browser", variable=self.show_browser,
            font=self._f(12), text_color=TEXT_MID,
            progress_color=CG_RED, button_color=SURFACE,
            button_hover_color=ELEVATED, fg_color=BORDER
        ).pack(side="left", padx=(0, S16))

        self.cred_lbl = ctk.CTkLabel(ctrl, text="",
                                     font=self._f(11), text_color=TEXT_DIM)
        self.cred_lbl.pack(side="left")

        self.run_btn = self._btn_primary(
            ctrl, "Run Automation", self._run_automation,
            height=36, width=140)
        self.run_btn.pack(side="right")

        # --- progress row ---
        prog_card = self._card(f)
        prog_card.pack(fill="x", pady=(0, S6))

        prog = ctk.CTkFrame(prog_card, fg_color="transparent")
        prog.pack(fill="x", padx=S16, pady=S10)

        top_row = ctk.CTkFrame(prog, fg_color="transparent")
        top_row.pack(fill="x", pady=(0, S6))

        self.status_lbl = ctk.CTkLabel(
            top_row, text="Ready", font=self._f(12),
            text_color=TEXT_MID, anchor="w")
        self.status_lbl.pack(side="left", fill="x", expand=True)

        self.progress_bar = ctk.CTkProgressBar(
            prog, height=8, corner_radius=4,
            fg_color=ELEVATED, progress_color=CG_RED)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        # --- banner slot ---
        self._banner_slot = ctk.CTkFrame(f, fg_color="transparent")
        self._banner_slot.pack(fill="x")

        # --- activity log (fills remaining space) ---
        log_card = self._card(f)
        log_card.pack(fill="both", expand=True, pady=(0, S6))

        lh = ctk.CTkFrame(log_card, fg_color="transparent")
        lh.pack(fill="x", padx=S16, pady=(S8, S4))
        ctk.CTkLabel(lh, text="ACTIVITY LOG",
                     font=self._f(10,"bold"),
                     text_color=TEXT_DIM).pack(side="left")

        self.log_text = ctk.CTkTextbox(
            log_card, wrap="word", state="disabled",
            corner_radius=6, font=self._fm(11),
            fg_color=BG, text_color=TEXT, border_width=0)
        self.log_text.pack(fill="both", expand=True, padx=S8, pady=(0, S8))

    # ========================================================= ACCOUNTS
    def _build_accounts(self):
        f = self._tab_accounts

        hdr = ctk.CTkFrame(f, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, S6))
        ctk.CTkLabel(hdr, text="Manage Accounts",
                     font=self._f(18,"bold"), text_color=TEXT,
                     anchor="w").pack(anchor="w")
        ctk.CTkLabel(hdr,
                     text="E-volve SecureAssess login credentials — AES-256 encrypted locally.",
                     font=self._f(12), text_color=TEXT_MID,
                     anchor="w").pack(anchor="w")

        # add form
        add_card = self._card(f)
        add_card.pack(fill="x", pady=(0, S6))

        add_inner = ctk.CTkFrame(add_card, fg_color="transparent")
        add_inner.pack(fill="x", padx=S16, pady=S12)

        ctk.CTkLabel(add_inner, text="ADD NEW ACCOUNT",
                     font=self._f(10,"bold"), text_color=TEXT_DIM, anchor="w"
                     ).pack(anchor="w", pady=(0, S8))

        form = ctk.CTkFrame(add_inner, fg_color="transparent")
        form.pack(fill="x")
        self._acc_user = self._entry(form, placeholder_text="Username")
        self._acc_user.pack(side="left", fill="x", expand=True, padx=(0, S8))
        self._acc_pass = self._entry(form, placeholder_text="Password", show="*")
        self._acc_pass.pack(side="left", fill="x", expand=True, padx=(0, S8))
        self._btn_primary(form, "Add", self._add_account,
                          height=40, width=90).pack(side="right")

        # saved accounts list
        list_card = self._card(f)
        list_card.pack(fill="both", expand=True, pady=(0, S6))

        lh = ctk.CTkFrame(list_card, fg_color="transparent")
        lh.pack(fill="x", padx=S16, pady=(S8, S4))
        ctk.CTkLabel(lh, text="SAVED ACCOUNTS",
                     font=self._f(10,"bold"), text_color=TEXT_DIM
                     ).pack(side="left")

        self._acc_scroll = ctk.CTkScrollableFrame(
            list_card, fg_color="transparent", corner_radius=0)
        self._acc_scroll.pack(fill="both", expand=True, padx=S8, pady=(0, S8))

    def _refresh_account_data(self):
        if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            self.cred_lbl.configure(text="No accounts")
            self.account_dd.configure(values=["All Accounts"])
            self.account_var.set("All Accounts")
            return
        try:
            creds = self.manager.list_credentials(master_password=self.master_password)
            n = len(creds)
            self.cred_lbl.configure(text=f"{n} account{'s' if n!=1 else ''}")
            names = [c.get("username","?") for c in creds]
            self.account_dd.configure(values=["All Accounts"] + names)
        except Exception:
            self.cred_lbl.configure(text="Error")

    def _refresh_accounts_list(self):
        for w in self._acc_scroll.winfo_children():
            w.destroy()
        try:
            creds = self.manager.list_credentials(master_password=self.master_password)
            if not creds:
                ctk.CTkLabel(self._acc_scroll,
                             text="No accounts saved yet. Add one above.",
                             font=self._f(12), text_color=TEXT_DIM
                             ).pack(pady=S20)
                return
            for cred in creds:
                uname = cred.get("username","?")
                row = ctk.CTkFrame(self._acc_scroll, fg_color=BG,
                                   corner_radius=8, height=48)
                row.pack(fill="x", pady=(0, S4))
                row.pack_propagate(False)
                ctk.CTkLabel(row, text=uname, font=self._f(13), text_color=TEXT
                             ).pack(side="left", padx=S16)
                rbtn = ctk.CTkButton(
                    row, text="Remove", width=80, height=32, corner_radius=8,
                    font=self._f(11), fg_color=DANGER_BG, hover_color="#FFCDD2",
                    text_color=DANGER, border_color=DANGER, border_width=1)
                rbtn.pack(side="right", padx=S12)
                self._bind_remove(uname, rbtn)
        except Exception as e:
            ctk.CTkLabel(self._acc_scroll, text=f"Error: {e}",
                         font=self._f(12), text_color=DANGER).pack(pady=S16)

    def _bind_remove(self, username, btn):
        state = {"confirmed": False}

        def reset():
            if btn.winfo_exists():
                btn.configure(text="Remove"); state["confirmed"] = False

        def on_click():
            if not state["confirmed"]:
                state["confirmed"] = True
                btn.configure(text="Confirm?")
                self.root.after(3000, reset)
            else:
                try:
                    self.manager.remove_credential(username,
                                                   master_password=self.master_password)
                    self._refresh_accounts_list()
                    self._refresh_account_data()
                    self._toast(f"Account '{username}' removed", "info")
                except Exception as e:
                    self._toast(f"Error: {e}", "error")

        btn.configure(command=on_click)

    def _add_account(self):
        u = self._acc_user.get().strip()
        p = self._acc_pass.get().strip()
        if not u or not p:
            self._toast("Both username and password are required.", "error"); return
        try:
            ok = self.manager.add_credential(u, p, master_password=self.master_password)
            if ok:
                self._acc_user.delete(0,"end")
                self._acc_pass.delete(0,"end")
                self._refresh_accounts_list()
                self._refresh_account_data()
                self._toast(f"Account '{u}' added.", "success")
            else:
                self._toast(f"Account '{u}' already exists.", "error")
        except Exception as e:
            self._toast(f"Error: {e}", "error")

    # ========================================================= FILES
    def _build_files(self):
        f = self._tab_files

        hdr = ctk.CTkFrame(f, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, S6))
        ctk.CTkLabel(hdr, text="Files", font=self._f(18,"bold"),
                     text_color=TEXT, anchor="w").pack(anchor="w")
        ctk.CTkLabel(hdr, text="Open Excel reports, PDF folders, and log files.",
                     font=self._f(12), text_color=TEXT_MID,
                     anchor="w").pack(anchor="w")

        self._files_scroll = ctk.CTkScrollableFrame(
            f, fg_color="transparent", corner_radius=0)
        self._files_scroll.pack(fill="both", expand=True, pady=(0, S6))

    def _refresh_files_list(self):
        for w in self._files_scroll.winfo_children():
            w.destroy()
        years = []
        try:
            for item in sorted(os.listdir(BASE_DIR), reverse=True):
                full = os.path.join(BASE_DIR, item)
                if os.path.isdir(full) and item.isdigit() and len(item) == 4:
                    years.append((item, full))
        except Exception:
            pass

        if not years:
            ctk.CTkLabel(self._files_scroll,
                         text="No data files yet. Run the automation first.",
                         font=self._f(12), text_color=TEXT_DIM
                         ).pack(pady=S24)
            return

        for year, path in years:
            card = self._card(self._files_scroll)
            card.pack(fill="x", pady=(0, S6))
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=S16, pady=S12)

            ctk.CTkLabel(inner, text=year, font=self._f(16,"bold"),
                         text_color=TEXT).pack(side="left")

            brow = ctk.CTkFrame(inner, fg_color="transparent")
            brow.pack(side="right")

            excel   = os.path.join(path, "exam_results.xlsx")
            reports = os.path.join(path, "reports")
            logs    = os.path.join(path, "logs")

            if os.path.exists(excel):
                self._btn_primary(brow, "Open Excel",
                                  lambda p=excel: self._open_path(p),
                                  height=32, width=100
                                  ).pack(side="left", padx=(0, S8))
            if os.path.exists(reports):
                self._btn_secondary(brow, "Reports",
                                    lambda p=reports: self._open_path(p),
                                    height=32, width=80
                                    ).pack(side="left", padx=(0, S8))
            if os.path.exists(logs):
                self._btn_secondary(brow, "Logs",
                                    lambda p=logs: self._open_path(p),
                                    height=32, width=70
                                    ).pack(side="left")

    def _open_path(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            self._toast(f"Could not open: {e}", "error")

    # ========================================================= HELP
    def _build_help(self):
        f = self._tab_help
        scroll = ctk.CTkScrollableFrame(f, fg_color="transparent", corner_radius=0)
        scroll.pack(fill="both", expand=True, pady=(0, S6))

        def section(title, body):
            ctk.CTkLabel(scroll, text=title, font=self._f(13,"bold"),
                         text_color=TEXT, anchor="w"
                         ).pack(anchor="w", padx=S4, pady=(S12, S2))
            ctk.CTkLabel(scroll, text=body, font=self._f(12),
                         text_color=TEXT_MID, anchor="w",
                         justify="left", wraplength=800
                         ).pack(anchor="w", padx=S4, pady=(0, S2))

        ctk.CTkLabel(scroll, text="How to Use", font=self._f(18,"bold"),
                     text_color=TEXT, anchor="w"
                     ).pack(anchor="w", padx=S4, pady=(S4, S4))
        ctk.CTkLabel(scroll,
                     text="This tool automates the download of exam results and PDF "
                          "reports from the City & Guilds E-volve SecureAssess platform. "
                          "It logs into your account(s), reads the results table, saves "
                          "data to Excel, and downloads PDF reports automatically.",
                     font=self._f(12), text_color=TEXT_MID, anchor="w",
                     justify="left", wraplength=800
                     ).pack(anchor="w", padx=S4, pady=(0, S8))

        self._divider(scroll).pack(fill="x", pady=S4)

        section("Getting Started",
                "1.  Set a master password on first launch.\n"
                "2.  Go to the Accounts tab and add your E-volve login(s).\n"
                "3.  Return to Dashboard and click Run Automation.\n"
                "4.  Results are saved to YYYY/exam_results.xlsx, organised by year.")
        section("Dashboard",
                "Select which account to run, or leave on 'All Accounts' to process "
                "every saved login in sequence. The progress bar and activity log show "
                "real-time status. Toggle 'Show browser' to watch the automation.")
        section("Accounts",
                "Add or remove your E-volve SecureAssess login credentials. "
                "All passwords are AES-256 encrypted using your master password "
                "and stored locally. Nothing is ever sent to an external server.")
        section("Files",
                "After running the automation, access your Excel spreadsheets, "
                "downloaded PDF reports, and log files. Files are organised by year.")

        self._divider(scroll).pack(fill="x", pady=S4)

        section("Tips",
                "- Close any open Excel files before running automation.\n"
                "- The tool only downloads results not already in your spreadsheet.\n"
                "- Date filters are applied automatically.\n"
                "- If a run fails, check the activity log for details.\n"
                "- Forgot your master password? Use the Reset option on the login screen.")

        self._divider(scroll).pack(fill="x", pady=S4)

        ctk.CTkLabel(scroll, text="Open Source", font=self._f(13,"bold"),
                     text_color=TEXT, anchor="w"
                     ).pack(anchor="w", padx=S4, pady=(S12, S2))
        ctk.CTkLabel(scroll,
                     text=f"Built and maintained by {AUTHOR} (@{AUTHOR_HANDLE}). "
                          "If you find this useful, your support is appreciated.",
                     font=self._f(12), text_color=TEXT_MID, anchor="w",
                     justify="left", wraplength=800
                     ).pack(anchor="w", padx=S4, pady=(0, S10))

        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.pack(anchor="w", padx=S4, pady=(0, S8))
        self._btn_primary(btn_row, "Support on Ko-fi",
                          lambda: webbrowser.open(KOFI_URL),
                          height=36, width=150, fg_color=AMBER,
                          hover_color="#BF360C"
                          ).pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "Report an Issue",
                            lambda: webbrowser.open(ISSUES_URL),
                            height=36, width=150
                            ).pack(side="left", padx=(0, S8))
        self._btn_secondary(btn_row, "View on GitHub",
                            lambda: webbrowser.open(REPO_URL),
                            height=36, width=150
                            ).pack(side="left")

        self._divider(scroll).pack(fill="x", pady=S4)

        ctk.CTkLabel(scroll, text="Disclaimer", font=self._f(13,"bold"),
                     text_color=TEXT, anchor="w"
                     ).pack(anchor="w", padx=S4, pady=(S12, S2))
        ctk.CTkLabel(scroll,
                     text=f"{APP_TITLE} {APP_SUBTITLE} {APP_VER}\n\n"
                          "This is an unofficial tool and is not affiliated with, "
                          "endorsed by, or associated with City & Guilds. E-volve and "
                          "SecureAssess are trademarks of The City and Guilds of London "
                          "Institute.",
                     font=self._f(11), text_color=TEXT_DIM, anchor="w",
                     justify="left", wraplength=800
                     ).pack(anchor="w", padx=S4, pady=(0, S16))

    # ========================================================= TOAST
    def _toast(self, message, kind="info"):
        for w in self._toast_slot.winfo_children():
            w.destroy()
        colours = {
            "success": (SUCCESS_BG, SUCCESS),
            "error":   (DANGER_BG,  DANGER),
            "info":    (CG_RED_LIGHT, CG_RED),
        }
        bg, fg = colours.get(kind, colours["info"])
        t = ctk.CTkFrame(self._toast_slot, fg_color=bg,
                         corner_radius=8, height=36)
        t.pack(fill="x", pady=(0, S6))
        t.pack_propagate(False)
        ctk.CTkLabel(t, text=message, font=self._f(12),
                     text_color=fg).pack(side="left", padx=S16)
        ctk.CTkButton(t, text="✕", width=24, height=24, corner_radius=6,
                      fg_color="transparent", hover_color=ELEVATED,
                      text_color=TEXT_DIM, font=self._f(11),
                      command=t.destroy).pack(side="right", padx=S8)
        self.root.after(5000, lambda: t.destroy() if t.winfo_exists() else None)

    # ========================================================= LOGGING
    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_status(self, text, colour=TEXT_MID):
        self.status_lbl.configure(text=text, text_color=colour)

    def _set_progress(self, val):
        self.progress_bar.set(max(0.0, min(1.0, val)))

    # ====================================================== AUTOMATION
    def _check_excel_locks(self):
        locked = []
        for p in glob.glob(os.path.join(BASE_DIR, "*", "exam_results.xlsx")):
            try:
                with open(p, "a"):
                    pass
            except (IOError, PermissionError):
                locked.append(os.path.basename(os.path.dirname(p)))
        return locked

    def _run_automation(self):
        if self.automation_thread and self.automation_thread.is_alive():
            self._toast("Automation is already running.", "error"); return
        try:
            creds = self.manager.list_credentials(master_password=self.master_password)
            if not creds:
                self._toast("No accounts configured. Add one in Accounts.", "error")
                return
        except Exception as e:
            self._toast(f"Error: {e}", "error"); return

        locked = self._check_excel_locks()
        if locked:
            self._toast(f"Close Excel first: {', '.join(locked)}/exam_results.xlsx", "error")
            return

        sel = self.account_var.get()
        selected = None if sel == "All Accounts" else sel
        self._total_accounts = 1 if selected else len(creds)
        self._done_accounts  = 0
        self._acct_progress  = 0.0

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0","end")
        self.log_text.configure(state="disabled")
        for w in self._banner_slot.winfo_children():
            w.destroy()

        self.run_btn.configure(state="disabled", text="Running…", fg_color=ELEVATED,
                               text_color=TEXT_MID)
        self._set_progress(0)
        self._set_status("Starting…", CG_RED)

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

            stats = EvolveAutomation(
                headless, self.master_password, selected_username).run()
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

    # ----------------------------------------------- progress parsing
    def _update_progress(self, msg):
        W = self._W
        if "Starting for account" in msg:
            self._acct_progress = 0.0; self._acct_cur_page = 0
            self._acct_pages = 1; self._push_progress()
            self._set_status(self._stat("Logging in…"), CG_RED)
        elif "Login submitted" in msg:
            self._acct_progress = W["login"]; self._push_progress()
            self._set_status(self._stat("Loading results…"), CG_RED)
        elif "Results tab opened" in msg:
            self._acct_progress = W["login"]+W["navigate"]; self._push_progress()
        elif "Switched to Results iframe" in msg:
            self._acct_progress = W["login"]+W["navigate"]+W["iframe"]
            self._push_progress()
        elif "Refresh button clicked" in msg:
            self._acct_progress = W["login"]+W["navigate"]+W["iframe"]+W["refresh"]
            self._push_progress()
            self._set_status(self._stat("Applying filters…"), CG_RED)
        elif "Date filter updated" in msg:
            self._acct_progress = (W["login"]+W["navigate"]+W["iframe"]+
                                   W["refresh"]+W["filter"])
            self._push_progress()
            self._set_status(self._stat("Loading data…"), CG_RED)
        elif "page(s) to scrape" in msg:
            m = re.search(r"Found (\d+) page", msg)
            if m:
                self._acct_pages = max(int(m.group(1)), 1)
            base = W["login"]+W["navigate"]+W["iframe"]+W["refresh"]+W["filter"]+W["hashes"]
            self._acct_progress = base; self._push_progress()
            self._set_status(self._stat(f"Scraping {self._acct_pages} page(s)…"), CG_RED)
        elif "Processing page" in msg:
            m = re.search(r"Processing page (\d+)/(\d+)", msg)
            if m:
                pg = int(m.group(1)); self._acct_cur_page = pg
                base = (W["login"]+W["navigate"]+W["iframe"]+
                        W["refresh"]+W["filter"]+W["hashes"])
                self._acct_progress = base + W["scrape"]*(pg/self._acct_pages)
                self._push_progress()
                self._set_status(self._stat(f"Page {pg}/{self._acct_pages}"), CG_RED)
        elif "Processing PDF for" in msg:
            base = (W["login"]+W["navigate"]+W["iframe"]+
                    W["refresh"]+W["filter"]+W["hashes"]+W["scrape"])
            self._acct_progress = min(base+W["pdfs"]*0.5, 0.95)
            self._push_progress()
            self._set_status(self._stat("Downloading PDFs…"), CG_RED)
        elif "All done for account" in msg:
            self._acct_progress = 0.95; self._push_progress()
        elif "Chrome closed" in msg:
            self._done_accounts += 1; self._acct_progress = 0.0
            self._push_progress()

    def _stat(self, detail):
        return f"Account {self._done_accounts+1}/{self._total_accounts}  |  {detail}"

    def _push_progress(self):
        if self._total_accounts == 0:
            return
        overall = (self._done_accounts + self._acct_progress) / self._total_accounts
        self._set_progress(overall)

    # ----------------------------------------------- completion / error
    def _on_complete(self, stats):
        self.run_btn.configure(state="normal", text="Run Automation",
                               fg_color=CG_RED, text_color="#FFFFFF")
        self._set_progress(1.0)
        self._set_status("Completed", SUCCESS)
        sep = "─" * 48
        self._log(f"\n{sep}\n  COMPLETED\n"
                  f"  Accounts : {stats.accounts_processed}\n"
                  f"  New rows : {stats.new_rows_added}\n"
                  f"  PDFs     : {stats.pdfs_downloaded}\n"
                  f"  Errors   : {stats.errors_encountered}\n{sep}")
        for w in self._banner_slot.winfo_children():
            w.destroy()
        txt = (f"Completed — {stats.accounts_processed} account(s), "
               f"{stats.new_rows_added} new rows, "
               f"{stats.pdfs_downloaded} PDFs, "
               f"{stats.errors_encountered} error(s)")
        banner = ctk.CTkFrame(self._banner_slot, fg_color=SUCCESS_BG,
                              corner_radius=8, height=36)
        banner.pack(fill="x", pady=(0, S6))
        banner.pack_propagate(False)
        ctk.CTkLabel(banner, text=txt, font=self._f(12),
                     text_color=SUCCESS).pack(side="left", padx=S16)
        ctk.CTkButton(banner, text="✕", width=24, height=24, corner_radius=6,
                      fg_color="transparent", hover_color=ELEVATED,
                      text_color=TEXT_DIM, font=self._f(11),
                      command=banner.destroy).pack(side="right", padx=S8)

    def _on_error(self, msg):
        self.run_btn.configure(state="normal", text="Run Automation",
                               fg_color=CG_RED, text_color="#FFFFFF")
        self._set_progress(0)
        self._set_status("Failed", DANGER)
        self._log(f"\nERROR: {msg}\n")
        self._toast(f"Automation failed: {msg}", "error")

    # ================================================= WINDOW MANAGEMENT
    def _on_close(self):
        if self.automation_thread and self.automation_thread.is_alive():
            if not messagebox.askyesno("Confirm", "Automation is running. Quit anyway?"):
                return
        self.root.destroy()

    # ================================================= ENTRY POINT
    def _maximize(self):
        """Maximize window — called via after() so the event loop is running."""
        try:
            self.root.state("zoomed")
        except Exception:
            w, h = 1280, 800
            sx = max((self.root.winfo_screenwidth()  - w) // 2, 0)
            sy = max((self.root.winfo_screenheight() - h) // 2, 0)
            self.root.geometry(f"{w}x{h}+{sx}+{sy}")

    def run(self):
        """Launch the application maximised."""
        self._show_lock_screen()
        self.root.after(50, self._maximize)
        self.root.mainloop()
