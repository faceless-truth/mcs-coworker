"""
MC & S Desktop Agent — Main Application
"""
import customtkinter as ctk
import threading
import time
import sys
import os
from datetime import datetime
from tkinter import messagebox
import tkinter as tk

try:
    import pystray
    from PIL import Image as PILImage
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

import config
from config import (
    init_db, get_setting, set_setting, get_rules, save_rule, delete_rule,
    get_staff, save_staff, delete_staff, get_recent_activity,
    get_links, save_link, delete_link,
    get_style_preferences, save_style_preferences,
    add_feedback_message, get_feedback_history, clear_feedback_history,
    add_lesson, get_active_lessons, delete_lesson, toggle_lesson,
)
from graph_client import GraphClient
from plugin_loader import PluginLoader

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BRAND_BLUE    = "#1565C0"
BRAND_DARK    = "#0D47A1"
ACCENT_GREEN  = "#2E7D32"
ACCENT_AMBER  = "#E65100"
BG_LIGHT      = "#F5F7FA"
CARD_BG       = "#FFFFFF"
TEXT_PRIMARY   = "#1A1A2E"
TEXT_MUTED     = "#6B7280"
SUCCESS_FG     = "#2E7D32"
DRAFT_BG       = "#E3F2FD"
DRAFT_FG       = "#1565C0"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_db()
        self.title("MC & S — Desktop Agent")
        self.geometry("1220x760")
        self.minsize(1000, 660)
        self.configure(fg_color=BG_LIGHT)

        self._loader = PluginLoader(log_callback=self._log)
        self._loader.on_run_complete(self._on_plugin_run_complete)
        self._graph: GraphClient | None = None
        self._session_actions = 0
        self._session_drafts  = 0

        # Check if first-run setup is needed
        if not get_setting("setup_completed"):
            self._show_setup_wizard()
        else:
            self._launch_main_ui()

    def _launch_main_ui(self):
        """Build and show the main application interface."""
        # Destroy wizard frame if it exists
        if hasattr(self, '_wizard_frame'):
            self._wizard_frame.destroy()

        self._tray_icon = None
        self._tray_thread = None

        self._build_header()
        self._build_nav()
        self._build_content_area()
        self._load_saved_settings()
        self._try_restore_session()
        self.after(200, self._initialise_plugins)
        self._show_page("dashboard")

        # Override window close to minimise to tray instead of quitting
        self.protocol("WM_DELETE_WINDOW", self._on_close_requested)

    # ────────────────────────────────────────────────────────────────────────
    # First-Run Setup Wizard
    # ────────────────────────────────────────────────────────────────────────

    def _show_setup_wizard(self):
        """Display a step-by-step setup wizard for first-time users."""
        self._wizard_step = 0
        self._wizard_frame = ctk.CTkFrame(self, fg_color=BG_LIGHT, corner_radius=0)
        self._wizard_frame.pack(fill="both", expand=True)

        # Header bar
        hdr = ctk.CTkFrame(self._wizard_frame, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="🏢  MC & S  Coworker — Setup",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)

        # Progress indicator
        self._wizard_progress_frame = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT, height=50)
        self._wizard_progress_frame.pack(fill="x", padx=60, pady=(20, 0))
        self._wizard_progress_frame.pack_propagate(False)
        self._wizard_step_labels = []
        steps = ["Welcome", "Sign In", "AI Key", "Ready"]
        for i, step_name in enumerate(steps):
            lbl = ctk.CTkLabel(self._wizard_progress_frame,
                               text=f"  {i+1}. {step_name}  ",
                               font=ctk.CTkFont(size=13, weight="bold" if i == 0 else "normal"),
                               text_color=BRAND_BLUE if i == 0 else TEXT_MUTED)
            lbl.pack(side="left", padx=12)
            self._wizard_step_labels.append(lbl)

        # Content area for wizard steps
        self._wizard_content = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT)
        self._wizard_content.pack(fill="both", expand=True, padx=60, pady=(10, 20))

        # Button row
        self._wizard_btn_frame = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT, height=60)
        self._wizard_btn_frame.pack(fill="x", padx=60, pady=(0, 30))
        self._wizard_btn_frame.pack_propagate(False)

        self._wizard_back_btn = ctk.CTkButton(
            self._wizard_btn_frame, text="← Back", width=120, height=42,
            fg_color="transparent", hover_color="#E3F2FD",
            text_color=BRAND_BLUE, border_width=1, border_color=BRAND_BLUE,
            font=ctk.CTkFont(size=14), command=self._wizard_back)
        self._wizard_back_btn.pack(side="left")

        self._wizard_next_btn = ctk.CTkButton(
            self._wizard_btn_frame, text="Get Started →", width=180, height=42,
            fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=14, weight="bold"), command=self._wizard_next)
        self._wizard_next_btn.pack(side="right")

        self._wizard_show_step(0)

    def _wizard_update_progress(self, step):
        for i, lbl in enumerate(self._wizard_step_labels):
            if i < step:
                lbl.configure(text_color=ACCENT_GREEN,
                              font=ctk.CTkFont(size=13, weight="normal"))
            elif i == step:
                lbl.configure(text_color=BRAND_BLUE,
                              font=ctk.CTkFont(size=13, weight="bold"))
            else:
                lbl.configure(text_color=TEXT_MUTED,
                              font=ctk.CTkFont(size=13, weight="normal"))

    def _wizard_show_step(self, step):
        self._wizard_step = step
        self._wizard_update_progress(step)

        # Clear content area
        for w in self._wizard_content.winfo_children():
            w.destroy()

        # Update button states
        if step == 0:
            self._wizard_back_btn.configure(state="disabled")
            self._wizard_next_btn.configure(text="Get Started →")
        elif step == 3:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="✓  Launch MC & S Coworker")
        else:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="Next →")

        if step == 0:
            self._wizard_step_welcome()
        elif step == 1:
            self._wizard_step_signin()
        elif step == 2:
            self._wizard_step_apikey()
        elif step == 3:
            self._wizard_step_ready()

    def _wizard_step_welcome(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Welcome to MC & S Coworker",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(30, 6))
        ctk.CTkLabel(card, text="Your AI-powered desktop assistant for email triage and automation.",
                     font=ctk.CTkFont(size=14), text_color=TEXT_MUTED).pack(pady=(0, 20))

        features = [
            ("📨", "Email Triage", "Automatically classifies incoming emails and drafts replies"),
            ("🧠", "Learning Memory", "Remembers your preferences and improves over time"),
            ("🔌", "Plugin System", "Extensible — add new automations as your needs grow"),
        ]
        for icon, title, desc in features:
            row = ctk.CTkFrame(card, fg_color="#F0F4FF", corner_radius=10)
            row.pack(fill="x", padx=60, pady=3)
            ctk.CTkLabel(row, text=icon, font=ctk.CTkFont(size=20)).pack(side="left", padx=(16, 10), pady=10)
            text_frame = ctk.CTkFrame(row, fg_color="transparent")
            text_frame.pack(side="left", fill="x", expand=True, pady=6)
            ctk.CTkLabel(text_frame, text=title,
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=TEXT_PRIMARY, anchor="w").pack(anchor="w")
            ctk.CTkLabel(text_frame, text=desc,
                         font=ctk.CTkFont(size=12),
                         text_color=TEXT_MUTED, anchor="w").pack(anchor="w")

        ctk.CTkLabel(card, text="Let's get you set up — it only takes 2 minutes.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED).pack(pady=(16, 20))

    def _wizard_step_signin(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=20)

        ctk.CTkLabel(card, text="🔐", font=ctk.CTkFont(size=40)).pack(pady=(30, 8))
        ctk.CTkLabel(card, text="Connect to Microsoft 365",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(0, 4))
        ctk.CTkLabel(card, text="Sign in with your MC & S email so the agent can read and draft emails on your behalf.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED,
                     wraplength=500).pack(pady=(0, 20))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(padx=80, fill="x")

        # Pre-fill Tenant ID and Client ID (hidden from user in a collapsed section)
        saved_tid = get_setting("ms_tenant_id")
        saved_cid = get_setting("ms_client_id")

        ctk.CTkLabel(form, text="Your Email Address",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_email = ctk.CTkEntry(form, height=40,
                                          font=ctk.CTkFont(size=14),
                                          placeholder_text="e.g. sarah@mcands.com.au")
        saved_email = get_setting("ms_account_email")
        if saved_email:
            self._wizard_email.insert(0, saved_email)
        self._wizard_email.pack(fill="x", pady=(0, 16))

        # Advanced section for Tenant/Client ID (collapsible)
        adv_toggle = ctk.CTkButton(form, text="▸ Advanced — Entra ID credentials",
                                    fg_color="transparent", hover_color="#E3F2FD",
                                    text_color=TEXT_MUTED, anchor="w",
                                    font=ctk.CTkFont(size=12),
                                    height=28, width=300)
        adv_toggle.pack(anchor="w", pady=(0, 4))

        adv_frame = ctk.CTkFrame(form, fg_color="#F8F9FA", corner_radius=8)
        adv_visible = [False]

        def toggle_advanced():
            if adv_visible[0]:
                adv_frame.pack_forget()
                adv_toggle.configure(text="▸ Advanced — Entra ID credentials")
            else:
                adv_frame.pack(fill="x", pady=(0, 12))
                adv_toggle.configure(text="▾ Advanced — Entra ID credentials")
            adv_visible[0] = not adv_visible[0]

        adv_toggle.configure(command=toggle_advanced)

        ctk.CTkLabel(adv_frame, text="Tenant ID",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=12, pady=(8, 2))
        self._wizard_tenant = ctk.CTkEntry(adv_frame, height=34, font=ctk.CTkFont(size=12))
        if saved_tid:
            self._wizard_tenant.insert(0, saved_tid)
        self._wizard_tenant.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(adv_frame, text="Client ID (App ID)",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=12, pady=(4, 2))
        self._wizard_client = ctk.CTkEntry(adv_frame, height=34, font=ctk.CTkFont(size=12))
        if saved_cid:
            self._wizard_client.insert(0, saved_cid)
        self._wizard_client.pack(fill="x", padx=12, pady=(0, 10))

        ctk.CTkLabel(adv_frame,
                     text="These are pre-filled for MC & S. Only change if you have a different Entra ID app.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                     wraplength=400).pack(anchor="w", padx=12, pady=(0, 8))

        # Sign in button
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(pady=(8, 4))
        self._wizard_signin_btn = ctk.CTkButton(
            btn_frame, text="🔐  Sign in to Microsoft 365",
            width=280, height=44, fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._wizard_do_signin)
        self._wizard_signin_btn.pack()

        self._wizard_signin_status = ctk.CTkLabel(
            card, text="", font=ctk.CTkFont(size=13), text_color=TEXT_MUTED)
        self._wizard_signin_status.pack(pady=(4, 20))

    def _wizard_do_signin(self):
        email = self._wizard_email.get().strip()
        tid = self._wizard_tenant.get().strip()
        cid = self._wizard_client.get().strip()

        if not email:
            messagebox.showerror("Missing Email", "Please enter your email address.")
            return
        if not tid or not cid:
            messagebox.showerror("Missing Credentials",
                                 "Tenant ID and Client ID are required. Click 'Advanced' to enter them.")
            return

        # Save settings
        set_setting("ms_tenant_id", tid)
        set_setting("ms_client_id", cid)
        set_setting("ms_account_email", email)

        # Build graph client and authenticate
        self._graph = GraphClient(tid, cid)
        self._wizard_signin_status.configure(text="Opening browser…", text_color=TEXT_MUTED)
        self._wizard_signin_btn.configure(state="disabled")

        def callback(success, error):
            if success:
                self.after(0, self._wizard_signin_success)
            else:
                self.after(0, lambda: self._wizard_signin_fail(str(error)))

        self._graph.authenticate(callback=callback)

    def _wizard_signin_success(self):
        self._wizard_signin_status.configure(
            text="✓  Signed in successfully!", text_color=ACCENT_GREEN)
        self._wizard_signin_btn.configure(state="normal",
                                           text="✓  Signed In", fg_color=ACCENT_GREEN)

    def _wizard_signin_fail(self, error):
        self._wizard_signin_status.configure(
            text=f"✗  {error}", text_color="#C62828")
        self._wizard_signin_btn.configure(state="normal")

    def _wizard_step_apikey(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=20)

        ctk.CTkLabel(card, text="🤖", font=ctk.CTkFont(size=40)).pack(pady=(30, 8))
        ctk.CTkLabel(card, text="Connect to Claude AI",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(0, 4))
        ctk.CTkLabel(card, text="Claude powers the email classification and smart drafting. "
                              "You need an API key from Anthropic.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED,
                     wraplength=500).pack(pady=(0, 20))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(padx=80, fill="x")

        ctk.CTkLabel(form, text="Anthropic API Key",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_apikey = ctk.CTkEntry(form, height=40,
                                           font=ctk.CTkFont(size=14),
                                           show="*",
                                           placeholder_text="sk-ant-...")
        saved_key = get_setting("anthropic_api_key")
        if saved_key:
            self._wizard_apikey.insert(0, saved_key)
        self._wizard_apikey.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(form, text="Get your key from console.anthropic.com → API Keys",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(anchor="w")

        ctk.CTkLabel(card, text="",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(expand=True)

        note = ctk.CTkFrame(card, fg_color="#FFF3E0", corner_radius=10)
        note.pack(fill="x", padx=60, pady=(0, 30))
        ctk.CTkLabel(note, text="💡  If your firm uses a shared API key, ask your administrator for it. "
                              "You can always change this later in Settings.",
                     font=ctk.CTkFont(size=12), text_color="#E65100",
                     wraplength=480).pack(padx=16, pady=12)

    def _wizard_step_ready(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=20)

        ctk.CTkLabel(card, text="🎉", font=ctk.CTkFont(size=48)).pack(pady=(40, 10))
        ctk.CTkLabel(card, text="You're All Set!",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(0, 8))

        email = get_setting("ms_account_email") or "your mailbox"
        ctk.CTkLabel(card, text=f"MC & S Coworker is ready to start working for you.",
                     font=ctk.CTkFont(size=14), text_color=TEXT_MUTED).pack(pady=(0, 24))

        summary_items = [
            ("📨", f"Monitoring: {email}"),
            ("✏️", "Mode: Draft — emails are drafted for your review before sending"),
            ("✍️", "Signature: Automatically pulled from your recent sent emails"),
            ("🧠", "Memory: Use the Memory tab to teach the agent your preferences"),
        ]
        for icon, text in summary_items:
            row = ctk.CTkFrame(card, fg_color="#E8F5E9", corner_radius=8)
            row.pack(fill="x", padx=80, pady=3)
            ctk.CTkLabel(row, text=f"{icon}  {text}",
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_PRIMARY, anchor="w").pack(anchor="w", padx=16, pady=10)

        ctk.CTkLabel(card, text="Click the button below to launch the dashboard and start the scheduler.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED).pack(pady=(24, 30))

    def _wizard_next(self):
        if self._wizard_step == 0:
            # Welcome → Sign In
            self._wizard_show_step(1)

        elif self._wizard_step == 1:
            # Sign In → save and move to API key
            email = self._wizard_email.get().strip()
            tid = self._wizard_tenant.get().strip()
            cid = self._wizard_client.get().strip()
            if not email:
                messagebox.showerror("Missing Email", "Please enter your email address before continuing.")
                return
            if not tid or not cid:
                messagebox.showerror("Missing Credentials",
                                     "Please enter Tenant ID and Client ID. Click 'Advanced' to expand.")
                return
            set_setting("ms_tenant_id", tid)
            set_setting("ms_client_id", cid)
            set_setting("ms_account_email", email)
            self._wizard_show_step(2)

        elif self._wizard_step == 2:
            # API Key → save and move to Ready
            key = self._wizard_apikey.get().strip()
            if not key:
                messagebox.showerror("Missing API Key",
                                     "Please enter your Anthropic API key to continue.")
                return
            set_setting("anthropic_api_key", key)
            self._wizard_show_step(3)

        elif self._wizard_step == 3:
            # Ready → mark complete and launch main UI
            set_setting("setup_completed", "1")
            self._launch_main_ui()

    def _wizard_back(self):
        if self._wizard_step > 0:
            self._wizard_show_step(self._wizard_step - 1)

    def _initialise_plugins(self):
        self._loader.discover()
        self._loader.load_all()
        self._refresh_plugins_page()
        self._log(f"🔌  {len(self._loader.get_plugins())} plugin(s) discovered.")

    # ────────────────────────────────────────────────────────────────────────
    # Layout
    # ────────────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="🏢  MC & S  Desktop Agent",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)
        self.auth_status_label = ctk.CTkLabel(hdr, text="● Not signed in",
                                              text_color="#FFA726", font=ctk.CTkFont(size=12))
        self.auth_status_label.pack(side="right", padx=20)
        self.scheduler_label = ctk.CTkLabel(hdr, text="⏸  Scheduler: Off",
                                            text_color="#CFD8DC", font=ctk.CTkFont(size=12))
        self.scheduler_label.pack(side="right", padx=12)

    def _build_nav(self):
        nav = ctk.CTkFrame(self, width=210, fg_color=BRAND_DARK, corner_radius=0)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        self._nav_btns = {}
        pages = [
            ("dashboard", "🏠  Dashboard"),
            ("plugins",   "🧩  Plugins"),
            ("rules",     "📨  Email Rules"),
            ("staff",     "👥  Staff & Notify"),
            ("memory",    "🧠  Memory"),
            ("settings",  "🔧  Settings"),
            ("activity",  "📋  Activity Log"),
        ]
        ctk.CTkLabel(nav, text="", height=10).pack()
        for key, label in pages:
            btn = ctk.CTkButton(nav, text=label, width=200, height=42,
                                fg_color="transparent", hover_color="#1565C0",
                                text_color="white", anchor="w", font=ctk.CTkFont(size=13),
                                command=lambda k=key: self._show_page(k))
            btn.pack(pady=2, padx=6)
            self._nav_btns[key] = btn

    def _build_content_area(self):
        self.content = ctk.CTkFrame(self, fg_color=BG_LIGHT, corner_radius=0)
        self.content.pack(side="left", fill="both", expand=True)
        self._pages = {}
        self._build_dashboard()
        self._build_plugins_page()
        self._build_rules_page()
        self._build_staff_page()
        self._build_memory_page()
        self._build_settings_page()
        self._build_activity_page()

    # ────────────────────────────────────────────────────────────────────────
    # Dashboard
    # ────────────────────────────────────────────────────────────────────────

    def _build_dashboard(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["dashboard"] = page

        ctk.CTkLabel(page, text="Dashboard",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 2))
        ctk.CTkLabel(page, text="Your MC & S desktop agent — automations running while you work.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 16))

        card_row = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        card_row.pack(fill="x", padx=28)
        self._stat_labels = {}
        for key, title, default, icon in [
            ("scheduler", "Scheduler",      "Stopped", "⏱"),
            ("plugins",   "Plugins",        "0",       "🧩"),
            ("actions",   "Actions Today",  "0",       "⚡"),
            ("drafts",    "Drafts Created", "0",       "📝"),
        ]:
            f = ctk.CTkFrame(card_row, fg_color=CARD_BG, corner_radius=12)
            f.pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkLabel(f, text=icon, font=ctk.CTkFont(size=28)).pack(pady=(16, 4))
            lbl = ctk.CTkLabel(f, text=default,
                               font=ctk.CTkFont(size=22, weight="bold"), text_color=TEXT_PRIMARY)
            lbl.pack()
            ctk.CTkLabel(f, text=title, text_color=TEXT_MUTED, font=ctk.CTkFont(size=12)).pack(pady=(0, 16))
            self._stat_labels[key] = lbl

        btn_row = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        btn_row.pack(fill="x", padx=28, pady=16)
        self.start_btn = ctk.CTkButton(btn_row, text="▶  Start Scheduler", width=190, height=44,
                                       fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                                       font=ctk.CTkFont(size=14, weight="bold"),
                                       command=self._start_scheduler)
        self.start_btn.pack(side="left", padx=(0, 10))
        self.stop_btn = ctk.CTkButton(btn_row, text="⏹  Stop Scheduler", width=190, height=44,
                                      fg_color="#C62828", hover_color="#7F0000",
                                      font=ctk.CTkFont(size=14, weight="bold"),
                                      state="disabled", command=self._stop_scheduler)
        self.stop_btn.pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="🧩  Manage Plugins", width=170, height=44,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK, font=ctk.CTkFont(size=14),
                      command=lambda: self._show_page("plugins")).pack(side="left")

        ctk.CTkLabel(page, text="Live Log", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28)
        self.log_box = ctk.CTkTextbox(page, height=300,
                                      font=ctk.CTkFont(family="Courier", size=12),
                                      fg_color="#1A1A2E", text_color="#E0E0E0", corner_radius=8)
        self.log_box.pack(fill="both", expand=True, padx=28, pady=(6, 20))
        self.log_box.configure(state="disabled")

    # ────────────────────────────────────────────────────────────────────────
    # Plugins page
    # ────────────────────────────────────────────────────────────────────────

    def _build_plugins_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["plugins"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Plugins",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="How to add a plugin →", width=170, height=32,
                      fg_color="transparent", hover_color="#E3F2FD",
                      text_color=BRAND_BLUE, border_width=1, border_color=BRAND_BLUE,
                      font=ctk.CTkFont(size=12),
                      command=self._show_plugin_help).pack(side="right")

        ctk.CTkLabel(page,
                     text="Each plugin is an independent automation. Toggle, configure, and run them individually.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 14))

        self.plugins_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self.plugins_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))

    def _refresh_plugins_page(self):
        for w in self.plugins_scroll.winfo_children():
            w.destroy()

        active   = [p for p in self._loader.get_plugins() if not p.is_template]
        template = [p for p in self._loader.get_plugins() if p.is_template]

        if active:
            ctk.CTkLabel(self.plugins_scroll, text="Installed Plugins",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT_MUTED).pack(anchor="w", pady=(4, 6))
            for lp in active:
                self._plugin_card(self.plugins_scroll, lp)

        if template:
            ctk.CTkLabel(self.plugins_scroll,
                         text="Templates — copy plugins/plugin_template.py to start a new one",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT_MUTED).pack(anchor="w", pady=(18, 6))
            for lp in template:
                self._plugin_template_card(self.plugins_scroll, lp)

        self._stat_labels["plugins"].configure(text=str(len(active)))

    def _plugin_card(self, parent, lp):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=12)
        card.pack(fill="x", pady=6)

        # Top row
        top = ctk.CTkFrame(card, fg_color=CARD_BG)
        top.pack(fill="x", padx=16, pady=(14, 4))
        ctk.CTkLabel(top, text=lp.icon, font=ctk.CTkFont(size=22)).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(top, text=lp.name,
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkLabel(top, text=f"v{lp.version}",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(side="left", padx=8)

        en_var = ctk.BooleanVar(value=lp.enabled)
        ctk.CTkSwitch(top, text="", variable=en_var, width=46,
                      command=lambda lp=lp, v=en_var: self._toggle_plugin_enabled(lp, v)).pack(side="right")
        ctk.CTkLabel(top, text="Enabled", text_color=TEXT_MUTED,
                     font=ctk.CTkFont(size=12)).pack(side="right", padx=4)

        ctk.CTkLabel(card, text=lp.description,
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
                     wraplength=800, anchor="w").pack(anchor="w", padx=16, pady=(0, 6))

        # Status strip
        status_row = ctk.CTkFrame(card, fg_color="#F8F9FA", corner_radius=8)
        status_row.pack(fill="x", padx=16, pady=(0, 8))

        def stat_col(lbl_text, val_text, color=TEXT_MUTED):
            f = ctk.CTkFrame(status_row, fg_color="transparent")
            f.pack(side="left", padx=16, pady=8)
            ctk.CTkLabel(f, text=lbl_text, text_color=TEXT_MUTED,
                         font=ctk.CTkFont(size=10)).pack()
            ctk.CTkLabel(f, text=val_text, text_color=color,
                         font=ctk.CTkFont(size=12, weight="bold")).pack()

        stat_col("STATUS", "Ready" if lp.is_ready else "Not Ready",
                 SUCCESS_FG if lp.is_ready else "#C62828")
        stat_col("SCHEDULE", lp.schedule_label)
        stat_col("LAST RUN", lp.last_run.strftime("%d/%m %H:%M") if lp.last_run else "Never")
        stat_col("RESULT", lp.last_result or "—")
        summary_short = (lp.last_summary[:40] + "…") if len(lp.last_summary) > 40 else (lp.last_summary or "—")
        stat_col("SUMMARY", summary_short)

        # Controls row
        ctrl = ctk.CTkFrame(card, fg_color=CARD_BG)
        ctrl.pack(fill="x", padx=16, pady=(0, 14))

        draft_var = ctk.BooleanVar(value=lp.draft_mode)
        ctk.CTkSwitch(ctrl, text="", variable=draft_var, width=46,
                      command=lambda lp=lp, v=draft_var: self._toggle_plugin_draft(lp, v)).pack(side="left")

        draft_lbl = ctk.CTkLabel(ctrl,
                                 text="✏️  Draft Mode" if lp.draft_mode else "⚡  Auto Mode",
                                 text_color=ACCENT_AMBER if lp.draft_mode else SUCCESS_FG,
                                 font=ctk.CTkFont(size=12))
        draft_lbl.pack(side="left", padx=(4, 20))

        def _update_draft_lbl(lp=lp, lbl=draft_lbl, v=draft_var):
            lbl.configure(
                text="✏️  Draft Mode" if v.get() else "⚡  Auto Mode",
                text_color=ACCENT_AMBER if v.get() else SUCCESS_FG
            )
        draft_var.trace_add("write", lambda *_: _update_draft_lbl())

        ctk.CTkLabel(ctrl, text="Schedule:", text_color=TEXT_MUTED,
                     font=ctk.CTkFont(size=12)).pack(side="left")

        schedule_map = {
            "Manual only": 0, "Every 1 min": 60, "Every 5 min": 300,
            "Every 15 min": 900, "Every 30 min": 1800,
            "Every 1 hr": 3600, "Every 4 hr": 14400, "Every 24 hr": 86400
        }
        current_label = min(schedule_map.items(),
                            key=lambda x: abs(x[1] - lp.schedule_seconds)
                            if (x[1] != 0 or lp.schedule_seconds == 0) else 9999)[0]
        sched_var = ctk.StringVar(value=current_label)
        ctk.CTkOptionMenu(ctrl, values=list(schedule_map.keys()), variable=sched_var,
                          width=140, height=30,
                          command=lambda val, lp=lp, m=schedule_map:
                              self._set_plugin_schedule(lp, m.get(val, 0))
                          ).pack(side="left", padx=8)

        ctk.CTkButton(ctrl, text="▶ Run Now", width=100, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      font=ctk.CTkFont(size=12),
                      command=lambda lp=lp: self._run_plugin_now(lp)).pack(side="right")

        # Plugin-specific settings
        schema = lp.instance.settings_schema()
        if schema:
            self._plugin_settings_panel(card, lp, schema)

    def _plugin_settings_panel(self, parent, lp, schema):
        frame = ctk.CTkFrame(parent, fg_color="#F0F4FF", corner_radius=8)
        frame.pack(fill="x", padx=16, pady=(0, 12))

        ctk.CTkLabel(frame, text="Plugin Settings",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=BRAND_BLUE).pack(anchor="w", padx=12, pady=(8, 4))

        entries = {}
        for field in schema:
            row = ctk.CTkFrame(frame, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=2)
            ctk.CTkLabel(row, text=field["label"], width=200,
                         text_color=TEXT_PRIMARY, font=ctk.CTkFont(size=12),
                         anchor="w").pack(side="left")

            current = lp.instance.get_plugin_setting(field["key"], field.get("default", ""))
            ftype = field.get("type", "text")

            if ftype == "bool":
                var = ctk.BooleanVar(value=current == "1")
                ctk.CTkCheckBox(row, text="", variable=var).pack(side="left")
                entries[field["key"]] = ("bool", var)
            else:
                show = "*" if ftype == "password" else ""
                e = ctk.CTkEntry(row, height=30, width=280,
                                 font=ctk.CTkFont(size=12), show=show)
                e.insert(0, current)
                e.pack(side="left")
                entries[field["key"]] = ("text", e)

            if field.get("help"):
                ctk.CTkLabel(row, text=field["help"],
                             text_color=TEXT_MUTED, font=ctk.CTkFont(size=10)).pack(side="left", padx=8)

        def save_plugin_settings():
            for key, (ftype, widget) in entries.items():
                val = ("1" if widget.get() else "0") if ftype == "bool" else widget.get().strip()
                lp.instance.set_plugin_setting(key, val)
            self._log(f"💾  Settings saved for {lp.name}.")

        ctk.CTkButton(frame, text="Save Plugin Settings", height=30, width=180,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=12),
                      command=save_plugin_settings).pack(anchor="w", padx=12, pady=(4, 10))

    def _plugin_template_card(self, parent, lp):
        card = ctk.CTkFrame(parent, fg_color="#F8F8F8", corner_radius=10,
                            border_width=1, border_color="#E0E0E0")
        card.pack(fill="x", pady=4)
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=10)
        ctk.CTkLabel(row, text=f"{lp.icon}  {lp.name}",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(row, text=lp.description, text_color=TEXT_MUTED,
                     font=ctk.CTkFont(size=12)).pack(side="left", padx=16)
        ctk.CTkLabel(row, text="Copy plugins/plugin_template.py to get started →",
                     text_color=BRAND_BLUE, font=ctk.CTkFont(size=11)).pack(side="right")

    def _show_plugin_help(self):
        win = ctk.CTkToplevel(self)
        win.title("How to Add a Plugin")
        win.geometry("640x500")
        win.grab_set()

        txt = ctk.CTkTextbox(win, font=ctk.CTkFont(family="Courier", size=12))
        txt.pack(fill="both", expand=True, padx=16, pady=16)
        txt.insert("1.0", """HOW TO ADD A NEW PLUGIN
═══════════════════════════════════════════════

Step 1: Copy the template
        Navigate to the plugins/ folder
        Duplicate plugin_template.py
        Rename it: plugin_your_name.py

Step 2: Edit the file
        Change the class name (e.g. class NOAWorkflowPlugin)
        Fill in: name, description, icon, detail
        Set: requires_graph, requires_claude (True/False)
        Set: default_schedule
        Declare plugin settings in settings_schema()
        Write your logic in run()

Step 3: Restart the app
        The Plugins tab will auto-discover your new plugin.

═══════════════════════════════════════════════
WHAT YOUR PLUGIN CAN DO

  context.graph.fetch_unread_emails(folder, n)
  context.graph.send_email(to, subject, body)
  context.graph.create_draft(to, subject, body)
  context.graph.flag_email(message_id)

  context.claude.messages.create(...)

  context.log("message") → dashboard log
  context.draft_mode → True/False
  self.get_plugin_setting("key")
  self.set_plugin_setting("key", value)
  self.log_activity(...)

═══════════════════════════════════════════════
PLUGIN IDEAS FOR MC & S

  NOA Workflow          detect ATO assessment notices
  FuseSign Monitor      nudge unsigned docs after X days
  Meeting Prep Brief    email client summary before appointments
  Debtor Follow-Up      automated overdue reminders
  ASIC Reminders        parse ASIC notices, create calendar entries
  Monthly Invoicing     surface retainer clients each month
  Client Check-Ins      prompt 6-monthly outreach to companies/trusts
""")
        txt.configure(state="disabled")

    def _toggle_plugin_enabled(self, lp, var):
        self._loader.set_plugin_enabled(lp.plugin_id, var.get())
        self._log(f"🔌  {lp.name}: {'enabled' if var.get() else 'disabled'}.")

    def _toggle_plugin_draft(self, lp, var):
        self._loader.set_plugin_draft_mode(lp.plugin_id, var.get())
        self._log(f"✏️  {lp.name}: {'Draft Mode' if var.get() else 'Auto-Send'}.")

    def _set_plugin_schedule(self, lp, seconds):
        self._loader.set_plugin_schedule(lp.plugin_id, seconds)
        self._log(f"⏱  {lp.name}: schedule → {lp.schedule_label}.")

    def _run_plugin_now(self, lp):
        if not self._graph or not self._graph.is_authenticated():
            messagebox.showwarning("Not Signed In", "Please sign in to Microsoft 365 in Settings first.")
            return

        def run():
            self._loader.set_graph(self._graph)
            self._loader.set_claude()
            self._loader.reload_plugin(lp.plugin_id)
            self._loader.run_plugin(lp.plugin_id, manual=True)
            self.after(0, self._refresh_plugins_page)

        threading.Thread(target=run, daemon=True).start()

    def _on_plugin_run_complete(self, plugin_id, result):
        self._session_actions += result.actions_taken
        self._session_drafts  += result.drafts_created
        self.after(0, self._update_dashboard_stats)

        # Send tray notification if drafts were created and window is hidden
        if result.drafts_created > 0 and not self.winfo_viewable():
            self.send_tray_notification(
                "MC & S Coworker",
                f"{result.drafts_created} new draft(s) ready for review in Outlook."
            )

    def _update_dashboard_stats(self):
        self._stat_labels["actions"].configure(text=str(self._session_actions))
        self._stat_labels["drafts"].configure(text=str(self._session_drafts))

    # ────────────────────────────────────────────────────────────────────────
    # Scheduler
    # ────────────────────────────────────────────────────────────────────────

    def _start_scheduler(self):
        if not self._graph or not self._graph.is_authenticated():
            messagebox.showwarning("Not Signed In", "Please sign in to Microsoft 365 first.")
            return
        if not get_setting("anthropic_api_key"):
            messagebox.showwarning("API Key Missing", "Please add your Anthropic API key in Settings.")
            return

        self._loader.set_graph(self._graph)
        self._loader.set_claude()
        self._loader.load_all()
        self._loader.start_scheduler()

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._stat_labels["scheduler"].configure(text="Running")
        self.scheduler_label.configure(text="▶ Scheduler: Running", text_color="#A5D6A7")
        self._log("▶ Scheduler started. Plugins running on their configured schedules.")
        self._refresh_plugins_page()

    def _stop_scheduler(self):
        self._loader.stop_scheduler()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self._stat_labels["scheduler"].configure(text="Stopped")
        self.scheduler_label.configure(text="⏸  Scheduler: Off", text_color="#CFD8DC")
        self._log("⏸ Scheduler stopped.")

    # ────────────────────────────────────────────────────────────────────────
    # Email Rules page
    # ────────────────────────────────────────────────────────────────────────

    def _build_rules_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["rules"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Email Rules",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="+ Add Rule", width=110, height=34,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=self._add_rule).pack(side="right")

        ctk.CTkLabel(page,
                     text="Used by the Email Triage plugin. Define categories, keywords, and response templates.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 12))

        self.rules_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self.rules_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_rules_list()

    def _refresh_rules_list(self):
        for w in self.rules_scroll.winfo_children():
            w.destroy()
        for rule in get_rules():
            self._rule_card(self.rules_scroll, rule)

    def _rule_card(self, parent, rule):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10)
        card.pack(fill="x", pady=6)

        top_row = ctk.CTkFrame(card, fg_color=CARD_BG)
        top_row.pack(fill="x", padx=16, pady=(12, 4))

        en_var = ctk.BooleanVar(value=bool(rule.get("enabled", 1)))
        ctk.CTkCheckBox(top_row, text="", variable=en_var, width=24,
                        command=lambda r=rule, v=en_var: self._toggle_rule_enabled(r, v)).pack(side="left")

        ctk.CTkLabel(top_row, text=rule["category"],
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=BRAND_BLUE).pack(side="left", padx=8)

        ctk.CTkButton(top_row, text="Edit", width=60, height=28,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=lambda r=rule: self._edit_rule(r)).pack(side="right")
        ctk.CTkButton(top_row, text="Delete", width=60, height=28,
                      fg_color="#C62828", hover_color="#7F0000",
                      command=lambda r=rule: self._delete_rule(r)).pack(side="right", padx=4)

        ctk.CTkLabel(card, text=f"Keywords: {rule['keywords']}",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
                     wraplength=700, anchor="w").pack(anchor="w", padx=16, pady=(0, 4))
        ctk.CTkLabel(card, text=f"Subject: {rule.get('subject_template','')[:80]}",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
                     anchor="w").pack(anchor="w", padx=16, pady=(0, 10))

    def _toggle_rule_enabled(self, rule, var):
        rule["enabled"] = 1 if var.get() else 0
        save_rule(rule)

    def _add_rule(self):
        self._rule_dialog(None)

    def _edit_rule(self, rule):
        self._rule_dialog(rule)

    def _delete_rule(self, rule):
        if messagebox.askyesno("Delete Rule", f"Delete rule '{rule['category']}'?"):
            delete_rule(rule["id"])
            self._refresh_rules_list()

    def _rule_dialog(self, rule=None):
        win = ctk.CTkToplevel(self)
        win.title("Edit Rule" if rule else "Add Rule")
        win.geometry("680x680")
        win.grab_set()

        fields = {}

        def row(label, default="", height=None):
            ctk.CTkLabel(win, text=label, font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(anchor="w", padx=24, pady=(12, 2))
            if height:
                w = ctk.CTkTextbox(win, height=height, font=ctk.CTkFont(size=12))
                w.insert("1.0", default)
            else:
                w = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13))
                w.insert(0, default)
            w.pack(fill="x", padx=24)
            return w

        fields["category"] = row("Category Name", rule["category"] if rule else "")
        fields["keywords"] = row("Keywords (comma separated)", rule["keywords"] if rule else "")
        fields["subject_template"] = row("Reply Subject Template",
                                         rule.get("subject_template", "") if rule else "")
        fields["body_template"] = row("Reply Body (HTML supported)",
                                      rule.get("body_template", "") if rule else "", height=240)

        ctk.CTkLabel(win, text="Use {client_name}, {date}, {subject} as placeholders.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=24)

        def save():
            bw = fields["body_template"]
            bval = bw.get("1.0", "end-1c") if isinstance(bw, ctk.CTkTextbox) else bw.get()
            r = {
                "id": rule["id"] if rule else None,
                "category": fields["category"].get().strip().upper().replace(" ", "_"),
                "keywords": fields["keywords"].get().strip(),
                "subject_template": fields["subject_template"].get().strip(),
                "body_template": bval,
                "enabled": rule.get("enabled", 1) if rule else 1,
                "sort_order": rule.get("sort_order", 99) if rule else 99,
            }
            save_rule(r)
            self._refresh_rules_list()
            win.destroy()

        ctk.CTkButton(win, text="Save Rule", height=42,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=save).pack(fill="x", padx=24, pady=16)

    # ────────────────────────────────────────────────────────────────────────
    # Staff page
    # ────────────────────────────────────────────────────────────────────────

    def _build_staff_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["staff"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Staff & Notifications",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="+ Add Staff", width=120, height=34,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=self._add_staff).pack(side="right")

        ctk.CTkLabel(page,
                     text="Staff listed here receive email notifications when a plugin creates a draft for review.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 16))

        info = ctk.CTkFrame(page, fg_color=DRAFT_BG, corner_radius=10)
        info.pack(fill="x", padx=28, pady=(0, 16))
        ctk.CTkLabel(info,
                     text="📝 Any plugin with Draft Mode ON will notify these staff members when a draft is ready in Outlook.",
                     text_color=DRAFT_FG, font=ctk.CTkFont(size=12),
                     wraplength=800).pack(padx=16, pady=12)

        self.staff_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self.staff_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_staff_list()

    def _refresh_staff_list(self):
        for w in self.staff_scroll.winfo_children():
            w.destroy()
        conn = config.get_db()
        rows = conn.execute("SELECT * FROM staff_notifications").fetchall()
        conn.close()
        for s in rows:
            self._staff_card(self.staff_scroll, dict(s))

    def _staff_card(self, parent, staff):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10)
        card.pack(fill="x", pady=5)
        row = ctk.CTkFrame(card, fg_color=CARD_BG)
        row.pack(fill="x", padx=16, pady=10)
        ctk.CTkLabel(row, text=f"👤 {staff['name']}",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkLabel(row, text=staff["email"],
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(side="left", padx=16)
        nv = ctk.BooleanVar(value=bool(staff.get("receives_drafts", 1)))
        ctk.CTkCheckBox(row, text="Receives draft notifications", variable=nv,
                        command=lambda s=staff, v=nv: self._toggle_staff_notify(s, v)).pack(side="left", padx=16)
        ctk.CTkButton(row, text="Delete", width=70, height=28,
                      fg_color="#C62828", hover_color="#7F0000",
                      command=lambda s=staff: self._del_staff(s)).pack(side="right")

    def _toggle_staff_notify(self, staff, var):
        staff["receives_drafts"] = 1 if var.get() else 0
        save_staff(staff)

    def _add_staff(self):
        win = ctk.CTkToplevel(self)
        win.title("Add Staff Member")
        win.geometry("440x260")
        win.grab_set()

        ctk.CTkLabel(win, text="Name", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=24, pady=(20, 2))
        ne = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13))
        ne.pack(fill="x", padx=24)

        ctk.CTkLabel(win, text="Email Address", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=24, pady=(12, 2))
        ee = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13))
        ee.pack(fill="x", padx=24)

        ctk.CTkButton(win, text="Save", height=40, fg_color=ACCENT_GREEN,
                      command=lambda: [save_staff({"name": ne.get().strip(), "email": ee.get().strip(),
                                                   "receives_drafts": 1, "enabled": 1}),
                                       self._refresh_staff_list(), win.destroy()]
                      ).pack(fill="x", padx=24, pady=16)

    def _del_staff(self, staff):
        if messagebox.askyesno("Delete", f"Remove {staff['name']}?"):
            delete_staff(staff["id"])
            self._refresh_staff_list()

    # ────────────────────────────────────────────────────────────────────────
    # Memory page
    # ────────────────────────────────────────────────────────────────────────

    def _build_memory_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["memory"] = page

        ctk.CTkLabel(page, text="Memory",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 2))
        ctk.CTkLabel(page,
                     text="Teach the agent how you like things done. Give feedback on drafts and it will learn for next time.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 12))

        # ── Main horizontal split ──
        body = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        body.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        # Left: Chat + Style
        left = ctk.CTkFrame(body, fg_color=BG_LIGHT)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))

        # ── Style Preferences Card ──
        style_card = ctk.CTkFrame(left, fg_color=CARD_BG, corner_radius=12)
        style_card.pack(fill="x", pady=(0, 10))

        style_top = ctk.CTkFrame(style_card, fg_color=CARD_BG)
        style_top.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(style_top, text="✍️  Tone & Style Preferences",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(style_top, text="Save", width=70, height=28,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=12),
                      command=self._save_style_prefs).pack(side="right")

        ctk.CTkLabel(style_card,
                     text="These instructions are included in every AI prompt. Write naturally.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=16, pady=(0, 4))

        self.style_textbox = ctk.CTkTextbox(style_card, height=80,
                                            font=ctk.CTkFont(size=12),
                                            fg_color="#F8F9FA", corner_radius=8)
        self.style_textbox.pack(fill="x", padx=16, pady=(0, 12))
        existing_style = get_style_preferences()
        if existing_style:
            self.style_textbox.insert("1.0", existing_style)

        # ── Chat Feedback Interface ──
        chat_card = ctk.CTkFrame(left, fg_color=CARD_BG, corner_radius=12)
        chat_card.pack(fill="both", expand=True)

        chat_top = ctk.CTkFrame(chat_card, fg_color=CARD_BG)
        chat_top.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(chat_top, text="💬  Feedback Chat",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(chat_top, text="Clear Chat", width=90, height=28,
                      fg_color="#C62828", hover_color="#7F0000",
                      font=ctk.CTkFont(size=11),
                      command=self._clear_memory_chat).pack(side="right")

        ctk.CTkLabel(chat_card,
                     text='Tell the agent what you liked or didn\'t like. e.g. "The Monday draft to John was too formal."',
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11),
                     wraplength=550).pack(anchor="w", padx=16, pady=(0, 6))

        self.chat_display = ctk.CTkTextbox(chat_card,
                                           font=ctk.CTkFont(size=12),
                                           fg_color="#1A1A2E", text_color="#E0E0E0",
                                           corner_radius=8, state="disabled")
        self.chat_display.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        # Configure tags for chat bubbles
        self.chat_display._textbox.tag_configure("user_name", foreground="#64B5F6",
                                                  font=("Arial", 11, "bold"))
        self.chat_display._textbox.tag_configure("agent_name", foreground="#81C784",
                                                  font=("Arial", 11, "bold"))
        self.chat_display._textbox.tag_configure("timestamp", foreground="#616161",
                                                  font=("Courier", 9))
        self.chat_display._textbox.tag_configure("user_msg", foreground="#E3F2FD",
                                                  font=("Arial", 12))
        self.chat_display._textbox.tag_configure("agent_msg", foreground="#C8E6C9",
                                                  font=("Arial", 12))

        input_row = ctk.CTkFrame(chat_card, fg_color=CARD_BG)
        input_row.pack(fill="x", padx=16, pady=(0, 12))

        self.chat_input = ctk.CTkEntry(input_row, height=40,
                                       placeholder_text="Type your feedback here...",
                                       font=ctk.CTkFont(size=13))
        self.chat_input.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.chat_input.bind("<Return>", lambda e: self._send_feedback())

        ctk.CTkButton(input_row, text="Send", width=80, height=40,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      font=ctk.CTkFont(size=13, weight="bold"),
                      command=self._send_feedback).pack(side="right")

        # ── Right side: Learned Lessons panel ──
        right = ctk.CTkFrame(body, fg_color=CARD_BG, corner_radius=12, width=320)
        right.pack(side="right", fill="y", padx=(8, 0))
        right.pack_propagate(False)

        ctk.CTkLabel(right, text="📚  Learned Lessons",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(right,
                     text="Extracted from your feedback. These guide every future email.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11),
                     wraplength=280).pack(anchor="w", padx=16, pady=(0, 8))

        self.lessons_scroll = ctk.CTkScrollableFrame(right, fg_color=CARD_BG)
        self.lessons_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 12))

        # Load existing data
        self._refresh_chat_display()
        self._refresh_lessons_panel()

    def _save_style_prefs(self):
        content = self.style_textbox.get("1.0", "end-1c")
        save_style_preferences(content)
        self._log("✍️  Style preferences saved.")
        messagebox.showinfo("Saved", "Tone & style preferences saved. These will be used in all future emails.")

    def _refresh_chat_display(self):
        self.chat_display.configure(state="normal")
        self.chat_display.delete("1.0", "end")
        history = get_feedback_history()
        if not history:
            self.chat_display.insert("end", "  No feedback yet.\n\n", "agent_msg")
            self.chat_display.insert("end",
                "  Start by telling me what you think of the emails I draft.\n"
                "  For example:\n"
                '  \u2022 "The email to John on Monday was too formal"\n'
                '  \u2022 "Always use Warm regards, not Kind regards"\n'
                '  \u2022 "Keep pricing emails under 3 paragraphs"\n',
                "agent_msg")
        else:
            for msg in history:
                ts = msg.get("timestamp", "")
                role = msg["role"]
                text = msg["message"]
                if role == "user":
                    self.chat_display.insert("end", f"  You", "user_name")
                    self.chat_display.insert("end", f"  {ts}\n", "timestamp")
                    self.chat_display.insert("end", f"  {text}\n\n", "user_msg")
                else:
                    self.chat_display.insert("end", f"  Agent", "agent_name")
                    self.chat_display.insert("end", f"  {ts}\n", "timestamp")
                    self.chat_display.insert("end", f"  {text}\n\n", "agent_msg")
        self.chat_display.see("end")
        self.chat_display.configure(state="disabled")

    def _send_feedback(self):
        text = self.chat_input.get().strip()
        if not text:
            return
        self.chat_input.delete(0, "end")

        # Save user message
        add_feedback_message("user", text)
        self._refresh_chat_display()

        # Process with Claude in background
        def process():
            api_key = get_setting("anthropic_api_key")
            if not api_key:
                agent_reply = (
                    "I\'ve noted your feedback, but I can\'t extract a lesson right now "
                    "because the Anthropic API key isn\'t configured. Please add it in Settings. "
                    "Your feedback has been saved and I\'ll process it once connected."
                )
                add_feedback_message("agent", agent_reply)
                # Still save the raw feedback as a lesson
                add_lesson(text, source="direct_feedback")
                self.after(0, self._refresh_chat_display)
                self.after(0, self._refresh_lessons_panel)
                return

            try:
                import anthropic
                client = anthropic.Anthropic(api_key=api_key)

                existing_lessons = get_active_lessons()
                lessons_context = ""
                if existing_lessons:
                    lessons_context = "\nExisting lessons already learned:\n" + "\n".join(
                        f"- {l['lesson']}" for l in existing_lessons
                    )

                prompt = f"""You are the memory system for MC & S Coworker, a desktop email agent for an accounting firm.

The user (Elio, the managing director) is giving you feedback about how the agent writes emails. Your job is to:
1. Acknowledge the feedback warmly and briefly
2. Extract a clear, actionable lesson that can be applied to future emails
3. Return your response in this exact JSON format:

{{"reply": "Your conversational acknowledgment", "lesson": "The clear, concise rule to remember"}}

Keep the reply friendly and under 2 sentences. The lesson should be a single, specific instruction.
If the feedback doesn't contain an actionable email-writing lesson, set lesson to null.
{lessons_context}

User feedback: {text}"""

                response = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=300,
                    messages=[{"role": "user", "content": prompt}],
                )

                import re, json
                raw = response.content[0].text.strip()
                raw = re.sub(r"```json\s*|```", "", raw).strip()
                parsed = json.loads(raw)

                agent_reply = parsed.get("reply", "Got it, I\'ll remember that.")
                lesson = parsed.get("lesson")

                add_feedback_message("agent", agent_reply)

                if lesson:
                    add_lesson(lesson, source=text[:100])
                    add_feedback_message("agent", f"📝 Lesson stored: \"{lesson}\"")

            except Exception as e:
                agent_reply = f"I\'ve saved your feedback. (Note: {e})"
                add_feedback_message("agent", agent_reply)
                add_lesson(text, source="direct_feedback")

            self.after(0, self._refresh_chat_display)
            self.after(0, self._refresh_lessons_panel)

        threading.Thread(target=process, daemon=True).start()

    def _clear_memory_chat(self):
        if messagebox.askyesno("Clear Chat", "Clear the feedback chat history?\n\nNote: Learned lessons will NOT be deleted."):
            clear_feedback_history()
            self._refresh_chat_display()

    def _refresh_lessons_panel(self):
        for w in self.lessons_scroll.winfo_children():
            w.destroy()

        lessons = get_active_lessons()
        if not lessons:
            ctk.CTkLabel(self.lessons_scroll,
                         text="No lessons yet. Start giving feedback in the chat!",
                         text_color=TEXT_MUTED, font=ctk.CTkFont(size=11),
                         wraplength=260).pack(pady=20)
            return

        for lesson in lessons:
            card = ctk.CTkFrame(self.lessons_scroll, fg_color="#F0F4FF", corner_radius=8)
            card.pack(fill="x", pady=3)

            top_row = ctk.CTkFrame(card, fg_color="transparent")
            top_row.pack(fill="x", padx=10, pady=(8, 2))

            ctk.CTkLabel(top_row, text=lesson["lesson"],
                         font=ctk.CTkFont(size=11),
                         text_color=TEXT_PRIMARY,
                         wraplength=220, anchor="w", justify="left").pack(side="left", fill="x", expand=True)

            ctk.CTkButton(top_row, text="✕", width=24, height=24,
                          fg_color="transparent", hover_color="#FFCDD2",
                          text_color="#C62828", font=ctk.CTkFont(size=12),
                          command=lambda lid=lesson["id"]: self._delete_lesson(lid)).pack(side="right")

            if lesson.get("source"):
                ctk.CTkLabel(card, text=f'From: "{lesson["source"]}"',
                             text_color=TEXT_MUTED, font=ctk.CTkFont(size=9),
                             wraplength=260, anchor="w").pack(anchor="w", padx=10, pady=(0, 6))

    def _delete_lesson(self, lesson_id):
        delete_lesson(lesson_id)
        self._refresh_lessons_panel()
        self._log("📚  Lesson removed from memory.")

    # ────────────────────────────────────────────────────────────────────────
    # Settings page
    # ────────────────────────────────────────────────────────────────────────

    def _build_settings_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["settings"] = page

        ctk.CTkLabel(page, text="Settings",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 4))

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        def section(title):
            ctk.CTkLabel(scroll, text=title, font=ctk.CTkFont(size=15, weight="bold"),
                         text_color=BRAND_BLUE).pack(anchor="w", pady=(20, 4))
            f = ctk.CTkFrame(scroll, fg_color=CARD_BG, corner_radius=10)
            f.pack(fill="x", pady=4)
            return f

        def field(parent, label, key, is_password=False):
            ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(10, 2))
            e = ctk.CTkEntry(parent, height=36, font=ctk.CTkFont(size=12),
                             show="*" if is_password else "")
            e.insert(0, get_setting(key))
            e.pack(fill="x", padx=16, pady=(0, 4))
            return e

        self._setting_entries = {}

        ms = section("Microsoft 365 / Entra ID")
        ctk.CTkLabel(ms, text="Register an app at portal.azure.com → Entra ID → App registrations. "
                              "Redirect URI: http://localhost:8765 (Public client/native).",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11),
                     wraplength=740).pack(anchor="w", padx=16, pady=(4, 0))
        self._setting_entries["ms_tenant_id"] = field(ms, "Tenant ID", "ms_tenant_id")
        self._setting_entries["ms_client_id"] = field(ms, "Client ID (App ID)", "ms_client_id")
        self._setting_entries["ms_account_email"] = field(ms, "Mailbox to Monitor", "ms_account_email")

        ar = ctk.CTkFrame(ms, fg_color=CARD_BG)
        ar.pack(fill="x", padx=16, pady=(0, 12))
        self.sign_in_btn = ctk.CTkButton(ar, text="🔑  Sign in to Microsoft 365",
                                         width=240, height=40,
                                         fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                                         command=self._sign_in)
        self.sign_in_btn.pack(side="left", pady=8)
        self.sign_in_status = ctk.CTkLabel(ar, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12))
        self.sign_in_status.pack(side="left", padx=12)

        ai = section("Claude AI (Anthropic)")
        ctk.CTkLabel(ai, text="Get your key from console.anthropic.com → API Keys.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=16, pady=(4, 0))
        self._setting_entries["anthropic_api_key"] = field(ai, "Anthropic API Key", "anthropic_api_key", True)

        pr = section("Practice Details")
        self._setting_entries["practice_name"] = field(pr, "Practice Name", "practice_name")
        self._setting_entries["monitor_folder"] = field(pr, "Default Folder to Watch", "monitor_folder")

        lf = section("Links & Forms")
        ctk.CTkLabel(lf,
                     text="Add links and forms here. Use the {tag} in any Email Rule template body to insert the URL.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11),
                     wraplength=740).pack(anchor="w", padx=16, pady=(4, 6))

        self._links_container = ctk.CTkFrame(lf, fg_color=CARD_BG)
        self._links_container.pack(fill="x", padx=16, pady=(0, 8))
        self._refresh_links_list()

        add_link_row = ctk.CTkFrame(lf, fg_color=CARD_BG)
        add_link_row.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(add_link_row, text="+ Add Link / Form", width=160, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      font=ctk.CTkFont(size=12),
                      command=self._add_link_dialog).pack(side="left", pady=6)

        bz = section("Business Hours")
        bi = ctk.CTkFrame(bz, fg_color=CARD_BG)
        bi.pack(fill="x", padx=16, pady=8)
        self.biz_enabled_var = ctk.BooleanVar(value=get_setting("business_hours_enabled", "1") == "1")
        ctk.CTkCheckBox(bi, text="Only run plugins during business hours",
                        variable=self.biz_enabled_var).pack(anchor="w", pady=4)
        self._setting_entries["business_hours_start"] = field(bz, "Start Hour (0-23)", "business_hours_start")
        self._setting_entries["business_hours_end"] = field(bz, "End Hour (0-23)", "business_hours_end")

        ctk.CTkButton(scroll, text="💾  Save All Settings", height=44,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._save_settings).pack(fill="x", pady=16)

    def _refresh_links_list(self):
        for w in self._links_container.winfo_children():
            w.destroy()

        links = get_links()
        if not links:
            ctk.CTkLabel(self._links_container,
                         text="No links yet. Click '+ Add Link / Form' to create one.",
                         text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(pady=12)
            return

        # Header
        hdr = ctk.CTkFrame(self._links_container, fg_color="transparent")
        hdr.pack(fill="x", padx=8, pady=(8, 2))
        ctk.CTkLabel(hdr, text="Name", width=180, font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(hdr, text="Tag (use in templates)", width=160, font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(hdr, text="URL", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=TEXT_MUTED).pack(side="left", padx=(0, 8))

        for link in links:
            row = ctk.CTkFrame(self._links_container, fg_color="#F0F4FF", corner_radius=6)
            row.pack(fill="x", padx=8, pady=2)

            ctk.CTkLabel(row, text=link["name"], width=180,
                         font=ctk.CTkFont(size=12),
                         text_color=TEXT_PRIMARY, anchor="w").pack(side="left", padx=(10, 4), pady=8)

            tag_label = f"{{{link['tag']}}}"
            ctk.CTkLabel(row, text=tag_label, width=160,
                         font=ctk.CTkFont(family="Courier", size=12),
                         text_color=BRAND_BLUE, anchor="w").pack(side="left", padx=4)

            url_text = link["url"] if link["url"] else "(not set — paste URL)"
            url_color = TEXT_PRIMARY if link["url"] else "#C62828"
            ctk.CTkLabel(row, text=url_text[:60] + ("..." if len(url_text) > 60 else ""),
                         font=ctk.CTkFont(size=11),
                         text_color=url_color, anchor="w").pack(side="left", fill="x", expand=True, padx=4)

            ctk.CTkButton(row, text="Edit", width=50, height=26,
                          fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                          font=ctk.CTkFont(size=11),
                          command=lambda l=link: self._edit_link_dialog(l)).pack(side="right", padx=2, pady=4)
            ctk.CTkButton(row, text="\u2715", width=30, height=26,
                          fg_color="transparent", hover_color="#FFCDD2",
                          text_color="#C62828", font=ctk.CTkFont(size=12),
                          command=lambda l=link: self._delete_link(l)).pack(side="right", padx=(0, 2), pady=4)

    def _add_link_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add Link / Form")
        dialog.geometry("500x280")
        dialog.resizable(False, False)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="Add a New Link or Form",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(16, 12))

        form = ctk.CTkFrame(dialog, fg_color="transparent")
        form.pack(fill="x", padx=24)

        ctk.CTkLabel(form, text="Name (e.g. Tax Return Checklist):",
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        name_entry = ctk.CTkEntry(form, height=32)
        name_entry.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(form, text="Tag (e.g. checklist_form — no spaces, lowercase):",
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        tag_entry = ctk.CTkEntry(form, height=32, placeholder_text="onboarding_form")
        tag_entry.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(form, text="URL (paste your Forms/link URL here):",
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        url_entry = ctk.CTkEntry(form, height=32, placeholder_text="https://forms.office.com/r/...")
        url_entry.pack(fill="x", pady=(2, 12))

        def save():
            name = name_entry.get().strip()
            tag = tag_entry.get().strip().lower().replace(" ", "_")
            url = url_entry.get().strip()
            if not name or not tag:
                messagebox.showerror("Missing Info", "Name and Tag are required.", parent=dialog)
                return
            save_link({"name": name, "tag": tag, "url": url, "enabled": 1})
            dialog.destroy()
            self._refresh_links_list()
            self._log(f"\U0001f517 Link added: {name} → {{{tag}}}")

        ctk.CTkButton(dialog, text="Save", height=36,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=save).pack(pady=(0, 16))

    def _edit_link_dialog(self, link):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Link / Form")
        dialog.geometry("500x280")
        dialog.resizable(False, False)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="Edit Link or Form",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(16, 12))

        form = ctk.CTkFrame(dialog, fg_color="transparent")
        form.pack(fill="x", padx=24)

        ctk.CTkLabel(form, text="Name:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        name_entry = ctk.CTkEntry(form, height=32)
        name_entry.pack(fill="x", pady=(2, 8))
        name_entry.insert(0, link["name"])

        ctk.CTkLabel(form, text=f"Tag:  {{{link['tag']}}}",
                     font=ctk.CTkFont(family="Courier", size=12),
                     text_color=TEXT_MUTED).pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(form, text="URL:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        url_entry = ctk.CTkEntry(form, height=32, placeholder_text="https://forms.office.com/r/...")
        url_entry.pack(fill="x", pady=(2, 12))
        if link["url"]:
            url_entry.insert(0, link["url"])

        def save():
            name = name_entry.get().strip()
            url = url_entry.get().strip()
            if not name:
                messagebox.showerror("Missing Info", "Name is required.", parent=dialog)
                return
            save_link({"id": link["id"], "name": name, "tag": link["tag"], "url": url, "enabled": 1})
            dialog.destroy()
            self._refresh_links_list()
            self._log(f"\U0001f517 Link updated: {name}")

        ctk.CTkButton(dialog, text="Save", height=36,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=save).pack(pady=(0, 16))

    def _delete_link(self, link):
        if messagebox.askyesno("Delete Link", f"Remove '{link['name']}'?\n\nMake sure no email rules are using {{{link['tag']}}} before deleting."):
            delete_link(link["id"])
            self._refresh_links_list()
            self._log(f"\U0001f517 Link removed: {link['name']}")

    def _save_settings(self):
        for key, entry in self._setting_entries.items():
            set_setting(key, entry.get().strip())
        set_setting("business_hours_enabled", "1" if self.biz_enabled_var.get() else "0")
        self._build_graph_client()
        messagebox.showinfo("Saved", "Settings saved.")

    # ────────────────────────────────────────────────────────────────────────
    # Activity log page
    # ────────────────────────────────────────────────────────────────────────

    def _build_activity_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["activity"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Activity Log",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="🔄 Refresh", width=100, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._refresh_activity).pack(side="right")

        self.activity_box = ctk.CTkTextbox(page,
                                           font=ctk.CTkFont(family="Courier", size=12),
                                           fg_color=CARD_BG, text_color=TEXT_PRIMARY, corner_radius=8)
        self.activity_box.pack(fill="both", expand=True, padx=28, pady=(12, 20))
        self._refresh_activity()

    def _refresh_activity(self):
        self.activity_box.configure(state="normal")
        self.activity_box.delete("1.0", "end")
        records = get_recent_activity(100)
        if not records:
            self.activity_box.insert("end", "No activity recorded yet.\n")
        else:
            self.activity_box.insert("end",
                f"{'Timestamp':<22} {'From':<35} {'Classification':<22} {'Action':<16} Draft\n")
            self.activity_box.insert("end", "─" * 110 + "\n")
            for r in records:
                self.activity_box.insert("end",
                    f"{r['timestamp']:<22} {str(r.get('from_email',''))[:33]:<35} "
                    f"{str(r.get('classification',''))[:20]:<22} "
                    f"{str(r.get('action',''))[:14]:<16} "
                    f"{'Yes' if r.get('draft_created') else 'No'}\n")
        self.activity_box.configure(state="disabled")

    # ────────────────────────────────────────────────────────────────────────
    # Navigation / Auth / Log
    # ────────────────────────────────────────────────────────────────────────

    def _show_page(self, key):
        for page in self._pages.values():
            page.pack_forget()
        self._pages[key].pack(fill="both", expand=True)
        for k, btn in self._nav_btns.items():
            btn.configure(fg_color="#1565C0" if k == key else "transparent")

    def _build_graph_client(self):
        tid = get_setting("ms_tenant_id")
        cid = get_setting("ms_client_id")
        if tid and cid:
            self._graph = GraphClient(tid, cid)
            return True
        return False

    def _try_restore_session(self):
        if self._build_graph_client() and self._graph.is_authenticated():
            self._on_auth_success()

    def _sign_in(self):
        self._save_settings()
        if not self._build_graph_client():
            messagebox.showerror("Missing Config", "Please enter Tenant ID and Client ID first.")
            return
        self.sign_in_status.configure(text="Opening browser…", text_color=TEXT_MUTED)
        self._graph.authenticate(callback=self._auth_callback)

    def _auth_callback(self, success, error):
        if success:
            self.after(0, self._on_auth_success)
        else:
            self.after(0, lambda: messagebox.showerror("Sign In Failed", str(error)))

    def _on_auth_success(self):
        self.auth_status_label.configure(text="● Signed in", text_color="#66BB6A")
        self.sign_in_status.configure(text="✓ Signed in", text_color=SUCCESS_FG)
        self._loader.set_graph(self._graph)
        self._loader.set_claude()
        self._log("🔑 Signed in to Microsoft 365.")

    # ────────────────────────────────────────────────────────────────────────
    # System Tray
    # ────────────────────────────────────────────────────────────────────────

    def _on_close_requested(self):
        """Called when the user clicks the X button. Minimise to tray if available."""
        if HAS_TRAY:
            self._minimise_to_tray()
        else:
            if messagebox.askyesno(
                "Close MC & S Coworker",
                "Closing the window will stop the scheduler.\n\n"
                "Are you sure you want to quit?"
            ):
                self._quit_app()

    def _minimise_to_tray(self):
        """Hide the window and show a system tray icon."""
        self.withdraw()  # Hide the window

        if self._tray_icon is not None:
            return  # Already in tray

        # Load the app icon for the tray
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon.png")
        if os.path.exists(icon_path):
            tray_image = PILImage.open(icon_path)
        else:
            # Fallback: create a simple blue square icon
            tray_image = PILImage.new("RGB", (64, 64), BRAND_BLUE)

        menu = pystray.Menu(
            pystray.MenuItem("Open MC & S Coworker", self._restore_from_tray, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Scheduler Running", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_from_tray),
        )

        self._tray_icon = pystray.Icon(
            "mcs_coworker",
            tray_image,
            "MC & S Coworker",
            menu,
        )

        self._tray_thread = threading.Thread(target=self._tray_icon.run, daemon=True)
        self._tray_thread.start()

    def _restore_from_tray(self, icon=None, item=None):
        """Restore the window from the system tray."""
        if self._tray_icon is not None:
            self._tray_icon.stop()
            self._tray_icon = None
            self._tray_thread = None

        self.after(0, self._do_restore)

    def _do_restore(self):
        """Restore window on the main thread."""
        self.deiconify()  # Show the window
        self.lift()       # Bring to front
        self.focus_force()

    def _quit_from_tray(self, icon=None, item=None):
        """Quit the app from the tray menu."""
        if self._tray_icon is not None:
            self._tray_icon.stop()
            self._tray_icon = None
        self.after(0, self._quit_app)

    def _quit_app(self):
        """Fully shut down the application."""
        try:
            if hasattr(self, '_loader') and self._loader:
                self._loader.stop()
        except Exception:
            pass
        self.destroy()

    def send_tray_notification(self, title: str, message: str):
        """Show a notification bubble from the system tray icon."""
        if self._tray_icon is not None and HAS_TRAY:
            try:
                self._tray_icon.notify(message, title)
            except Exception:
                pass  # Not all platforms support notifications

    def _log(self, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {message}\n"

        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", line)
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.after(0, _write)

    def _load_saved_settings(self):
        self._session_actions = 0
        self._session_drafts  = 0


if __name__ == "__main__":
    app = App()
    app.mainloop()
