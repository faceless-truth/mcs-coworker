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
    get_recent_activity,
    get_links, save_link, delete_link,
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
        self.title("MC & S — CoWorker")
        self.geometry("1220x760")
        self.minsize(1000, 660)
        self.configure(fg_color=BG_LIGHT)

        self._loader = PluginLoader(log_callback=self._log)
        self._loader.on_run_complete(self._on_plugin_run_complete)
        self._graph: GraphClient | None = None
        self._session_actions = 0
        self._session_drafts  = 0

        # Check if first-run setup is needed
        if get_setting("user_setup_complete") == "1":
            self._launch_main_ui()
        else:
            self._show_setup_wizard()

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
        self._try_restore_session()
        self.after(200, self._initialise_plugins)
        self._show_page("dashboard")

        # Auto-start scheduler after a short delay to allow plugin init
        self.after(1500, self._auto_start_scheduler)

        # Override window close to minimise to tray instead of quitting
        self.protocol("WM_DELETE_WINDOW", self._on_close_requested)

    # ────────────────────────────────────────────────────────────────────────
    # First-Run Onboarding Wizard (3 steps)
    # ────────────────────────────────────────────────────────────────────────

    def _show_setup_wizard(self):
        """Display a 3-step onboarding wizard for first-time users."""
        self._wizard_step = 0
        self._wizard_frame = ctk.CTkFrame(self, fg_color=BG_LIGHT, corner_radius=0)
        self._wizard_frame.pack(fill="both", expand=True)

        # Header bar
        hdr = ctk.CTkFrame(self._wizard_frame, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="MC & S  CoWorker",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)

        # Progress indicator
        self._wizard_progress_frame = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT, height=50)
        self._wizard_progress_frame.pack(fill="x", padx=60, pady=(20, 0))
        self._wizard_progress_frame.pack_propagate(False)
        self._wizard_step_labels = []
        steps = ["Welcome", "Your Details", "Connect Email"]
        for i, step_name in enumerate(steps):
            lbl = ctk.CTkLabel(self._wizard_progress_frame,
                               text=f"  Step {i+1}: {step_name}  ",
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
            self._wizard_btn_frame, text="Back", width=120, height=42,
            fg_color="transparent", hover_color="#E3F2FD",
            text_color=BRAND_BLUE, border_width=1, border_color=BRAND_BLUE,
            font=ctk.CTkFont(size=14), command=self._wizard_back)
        self._wizard_back_btn.pack(side="left")

        self._wizard_next_btn = ctk.CTkButton(
            self._wizard_btn_frame, text="Get Started", width=180, height=42,
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
            self._wizard_next_btn.configure(text="Get Started", state="normal")
        elif step == 1:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="Next", state="normal")
        elif step == 2:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="Finish Setup", state="disabled")

        if step == 0:
            self._wizard_step_welcome()
        elif step == 1:
            self._wizard_step_details()
        elif step == 2:
            self._wizard_step_connect()

    def _wizard_step_welcome(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Welcome to MC&S CoWorker",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(50, 8))
        ctk.CTkLabel(card, text="Let's get you set up in 3 easy steps.",
                     font=ctk.CTkFont(size=14), text_color=TEXT_MUTED).pack(pady=(0, 30))

        features = [
            ("📨", "Email Triage", "Automatically classifies incoming emails and drafts replies"),
            ("✏️", "Draft Mode", "Emails are drafted for your review before sending"),
            ("⚡", "Always On", "Runs in the background while you work"),
        ]
        for icon, title, desc in features:
            row = ctk.CTkFrame(card, fg_color="#F0F4FF", corner_radius=10)
            row.pack(fill="x", padx=80, pady=3)
            ctk.CTkLabel(row, text=icon, font=ctk.CTkFont(size=20)).pack(side="left", padx=(16, 10), pady=10)
            text_frame = ctk.CTkFrame(row, fg_color="transparent")
            text_frame.pack(side="left", fill="x", expand=True, pady=6)
            ctk.CTkLabel(text_frame, text=title,
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=TEXT_PRIMARY, anchor="w").pack(anchor="w")
            ctk.CTkLabel(text_frame, text=desc,
                         font=ctk.CTkFont(size=12),
                         text_color=TEXT_MUTED, anchor="w").pack(anchor="w")

    def _wizard_step_details(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Your Details",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(40, 4))
        ctk.CTkLabel(card, text="Tell us a bit about you so we can personalise your experience.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED).pack(pady=(0, 24))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(padx=100, fill="x")

        ctk.CTkLabel(form, text="Your Name",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_name = ctk.CTkEntry(form, height=40,
                                         font=ctk.CTkFont(size=14),
                                         placeholder_text="e.g. Sarah Chen")
        saved_name = get_setting("user_name")
        if saved_name:
            self._wizard_name.insert(0, saved_name)
        self._wizard_name.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(form, text="Your Firm Name",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_firm = ctk.CTkEntry(form, height=40,
                                         font=ctk.CTkFont(size=14),
                                         placeholder_text="e.g. Smith & Associates Accounting")
        saved_firm = get_setting("user_firm")
        if saved_firm:
            self._wizard_firm.insert(0, saved_firm)
        self._wizard_firm.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(form, text="Your Email Address",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_email = ctk.CTkEntry(form, height=40,
                                          font=ctk.CTkFont(size=14),
                                          placeholder_text="e.g. sarah@smithaccounting.com.au")
        saved_email = get_setting("user_email")
        if saved_email:
            self._wizard_email.insert(0, saved_email)
        self._wizard_email.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(form, text="This is the mailbox CoWorker will monitor for incoming emails.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(anchor="w")

    def _wizard_step_connect(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Connect your Microsoft 365 account",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(50, 8))
        ctk.CTkLabel(card, text="Click below to sign in. We'll only access your inbox to help\nyou respond to emails.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED,
                     wraplength=500, justify="center").pack(pady=(0, 30))

        self._wizard_signin_btn = ctk.CTkButton(
            card, text="Sign in to Microsoft 365",
            width=300, height=48, fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._wizard_do_signin)
        self._wizard_signin_btn.pack()

        # Spinner / status area
        self._wizard_signin_status = ctk.CTkLabel(
            card, text="", font=ctk.CTkFont(size=13), text_color=TEXT_MUTED)
        self._wizard_signin_status.pack(pady=(12, 0))

        # Success area (hidden initially)
        self._wizard_success_frame = ctk.CTkFrame(card, fg_color="#E8F5E9", corner_radius=10)

    def _wizard_do_signin(self):
        self._graph = GraphClient()
        self._wizard_signin_status.configure(text="Opening browser... please sign in.",
                                             text_color=TEXT_MUTED)
        self._wizard_signin_btn.configure(state="disabled")

        def callback(success, error):
            if success:
                self.after(0, self._wizard_signin_success)
            else:
                self.after(0, lambda: self._wizard_signin_fail(str(error)))

        self._graph.authenticate(callback=callback)

    def _wizard_signin_success(self):
        # Get user info
        email = get_setting("user_email") or ""
        try:
            info = self._graph.get_user_info()
            email = info.get("mail", info.get("userPrincipalName", email))
        except Exception:
            pass

        self._wizard_signin_status.configure(text="", text_color=TEXT_MUTED)
        self._wizard_signin_btn.configure(
            text="Connected", fg_color=ACCENT_GREEN, state="disabled")

        # Show success indicator
        self._wizard_success_frame.pack(fill="x", padx=80, pady=(16, 0))
        ctk.CTkLabel(self._wizard_success_frame,
                     text=f"Connected as {email}",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=ACCENT_GREEN).pack(padx=20, pady=12)

        # Enable the Finish Setup button
        self._wizard_next_btn.configure(state="normal")

    def _wizard_signin_fail(self, error):
        self._wizard_signin_status.configure(
            text=f"Sign in failed: {error}", text_color="#C62828")
        self._wizard_signin_btn.configure(state="normal")

    def _wizard_next(self):
        if self._wizard_step == 0:
            self._wizard_show_step(1)

        elif self._wizard_step == 1:
            name = self._wizard_name.get().strip()
            firm = self._wizard_firm.get().strip()
            email = self._wizard_email.get().strip()
            if not name:
                messagebox.showerror("Required", "Please enter your name.")
                return
            if not firm:
                messagebox.showerror("Required", "Please enter your firm name.")
                return
            if not email:
                messagebox.showerror("Required", "Please enter your email address.")
                return
            # Save details
            set_setting("user_name", name)
            set_setting("user_firm", firm)
            set_setting("user_email", email)
            set_setting("ms_account_email", email)
            set_setting("practice_name", firm)
            self._wizard_show_step(2)

        elif self._wizard_step == 2:
            # Finish — mark setup complete and launch
            set_setting("user_setup_complete", "1")
            self._launch_main_ui()

    def _wizard_back(self):
        if self._wizard_step > 0:
            self._wizard_show_step(self._wizard_step - 1)

    def _initialise_plugins(self):
        self._loader.discover()
        self._loader.load_all()

    # ────────────────────────────────────────────────────────────────────────
    # Layout
    # ────────────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="MC & S  CoWorker",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)

        # Auth status
        self.auth_status_label = ctk.CTkLabel(hdr, text="",
                                              text_color="#66BB6A", font=ctk.CTkFont(size=12))
        self.auth_status_label.pack(side="right", padx=20)

        # User name
        user_name = get_setting("user_name")
        if user_name:
            ctk.CTkLabel(hdr, text=user_name,
                         text_color="white", font=ctk.CTkFont(size=13)).pack(side="right", padx=8)

        # Scheduler status
        self.scheduler_label = ctk.CTkLabel(hdr, text="",
                                            text_color="#CFD8DC", font=ctk.CTkFont(size=12))
        self.scheduler_label.pack(side="right", padx=12)

    def _build_nav(self):
        nav = ctk.CTkFrame(self, width=210, fg_color=BRAND_DARK, corner_radius=0)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        self._nav_btns = {}
        pages = [
            ("dashboard", "Dashboard"),
            ("rules",     "Email Rules"),
            ("activity",  "Activity Log"),
        ]
        ctk.CTkLabel(nav, text="", height=10).pack()
        for key, label in pages:
            btn = ctk.CTkButton(nav, text=f"  {label}", width=200, height=42,
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
        self._build_rules_page()
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
        ctk.CTkLabel(page, text="Your MC & S CoWorker — automating email triage while you work.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 16))

        card_row = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        card_row.pack(fill="x", padx=28)
        self._stat_labels = {}
        for key, title, default, icon in [
            ("scheduler", "Status",           "Starting...", "⏱"),
            ("actions",   "Emails Processed",  "0",           "⚡"),
            ("drafts",    "Drafts Created",    "0",           "📝"),
        ]:
            f = ctk.CTkFrame(card_row, fg_color=CARD_BG, corner_radius=12)
            f.pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkLabel(f, text=icon, font=ctk.CTkFont(size=28)).pack(pady=(16, 4))
            lbl = ctk.CTkLabel(f, text=default,
                               font=ctk.CTkFont(size=22, weight="bold"), text_color=TEXT_PRIMARY)
            lbl.pack()
            ctk.CTkLabel(f, text=title, text_color=TEXT_MUTED, font=ctk.CTkFont(size=12)).pack(pady=(0, 16))
            self._stat_labels[key] = lbl

        ctk.CTkLabel(page, text="Live Log", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(16, 0))
        self.log_box = ctk.CTkTextbox(page, height=300,
                                      font=ctk.CTkFont(family="Courier", size=12),
                                      fg_color="#1A1A2E", text_color="#E0E0E0", corner_radius=8)
        self.log_box.pack(fill="both", expand=True, padx=28, pady=(6, 20))
        self.log_box.configure(state="disabled")

    # ────────────────────────────────────────────────────────────────────────
    # Scheduler (auto-start)
    # ────────────────────────────────────────────────────────────────────────

    def _auto_start_scheduler(self):
        """Auto-start the scheduler silently after setup is complete."""
        if not self._graph or not self._graph.is_authenticated():
            self._log("Waiting for Microsoft 365 connection...")
            self.after(3000, self._auto_start_scheduler)
            return

        self._loader.set_graph(self._graph)
        self._loader.set_claude()
        self._loader.load_all()
        self._loader.start_scheduler()

        self._stat_labels["scheduler"].configure(text="Running")
        self.scheduler_label.configure(text="Running", text_color="#A5D6A7")
        self._log("Scheduler started. Email triage running automatically.")

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
                     text="Define categories, keywords, and response templates for automatic email replies.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 12))

        self.rules_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self.rules_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_rules_list()

    def _refresh_rules_list(self):
        for w in self.rules_scroll.winfo_children():
            w.destroy()

        rules = get_rules()
        if not rules:
            # Empty state
            empty = ctk.CTkFrame(self.rules_scroll, fg_color=CARD_BG, corner_radius=12)
            empty.pack(fill="x", pady=20)
            ctk.CTkLabel(empty, text="No rules yet.",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(pady=(30, 4))
            ctk.CTkLabel(empty, text="Click 'Add Rule' to create your first auto-response.",
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_MUTED).pack(pady=(0, 30))
            return

        for rule in rules:
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
        ctk.CTkButton(top, text="Refresh", width=100, height=32,
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
            self.activity_box.insert("end", "-" * 110 + "\n")
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

    def _try_restore_session(self):
        # If already authenticated (e.g. from setup wizard), just update the UI
        if self._graph and self._graph.is_authenticated():
            self._on_auth_success()
            return
        # Otherwise try to restore from hardcoded credentials
        self._graph = GraphClient()
        if self._graph.is_authenticated():
            self._on_auth_success()

    def _on_auth_success(self):
        self.auth_status_label.configure(text="Connected", text_color="#66BB6A")
        self._loader.set_graph(self._graph)
        self._log("Connected to Microsoft 365.")

    def _on_plugin_run_complete(self, plugin_id, result):
        self._session_actions += result.actions_taken
        self._session_drafts  += result.drafts_created
        self.after(0, self._update_dashboard_stats)

        # Send tray notification if drafts were created and window is hidden
        if result.drafts_created > 0 and not self.winfo_viewable():
            self.send_tray_notification(
                "MC & S CoWorker",
                f"{result.drafts_created} new draft(s) ready for review in Outlook."
            )

    def _update_dashboard_stats(self):
        self._stat_labels["actions"].configure(text=str(self._session_actions))
        self._stat_labels["drafts"].configure(text=str(self._session_drafts))

    # ────────────────────────────────────────────────────────────────────────
    # System Tray
    # ────────────────────────────────────────────────────────────────────────

    def _on_close_requested(self):
        """Called when the user clicks the X button. Minimise to tray if available."""
        if HAS_TRAY:
            self._minimise_to_tray()
        else:
            if messagebox.askyesno(
                "Close MC & S CoWorker",
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
            pystray.MenuItem("Open MC & S CoWorker", self._restore_from_tray, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Scheduler Running", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_from_tray),
        )

        self._tray_icon = pystray.Icon(
            "mcs_coworker",
            tray_image,
            "MC & S CoWorker",
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
                self._loader.stop_scheduler()
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


if __name__ == "__main__":
    app = App()
    app.mainloop()
