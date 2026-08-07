"""
📧 Email Automation
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bulk email sender with attachments, Excel contacts & email templates.

Libraries:  smtplib (built-in) + pandas (Excel/CSV reading)
Run:        python email_automation.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
import os
import re
import smtplib
import ssl
import threading
from email.message import EmailMessage
from email.utils import formataddr
from tkinter import filedialog, messagebox

import pandas as pd
import tkinter as tk
from tkinter import ttk

# ── Palette (matches other DAA apps) ────────────────────────────
BG      = "#0f0f1a"
SIDEBAR = "#12121f"
CARD    = "#1a1a2e"
CARD2   = "#16213e"
BORDER  = "#2a3045"
TEXT    = "#f1f5f9"
SUBTEXT = "#64748b"
ACCENT  = "#a855f7"   # purple accent
BLUE    = "#3b82f6"
GREEN   = "#22c55e"
AMBER   = "#f59e0b"
RED     = "#ef4444"
WHITE   = "#ffffff"

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SMALL = ("Segoe UI", 8)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "email_templates")


def render_template(text, row):
    """Replace {{placeholder}} tokens with values from an Excel row dict."""
    def repl(m):
        key = m.group(1).strip()
        return str(row.get(key, "")) if key in row else m.group(0)
    return re.sub(r"\{\{\s*(\w+)\s*\}\}", repl, text)


class EmailAutomation:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Email Automation")
        self.root.geometry("980x720")
        self.root.configure(bg=BG)
        self.root.minsize(820, 620)

        self.contacts = pd.DataFrame()
        self.attachments = []
        self.templates = {}
        self._load_templates()

        self._build_ui()
        self._log("Ready. Configure SMTP, load contacts & template, then send.")

    # ── Template loading ─────────────────────────────────────────
    def _load_templates(self):
        os.makedirs(TEMPLATE_DIR, exist_ok=True)
        for f in sorted(os.listdir(TEMPLATE_DIR)):
            if f.endswith(".txt"):
                path = os.path.join(TEMPLATE_DIR, f)
                with open(path, "r", encoding="utf-8") as fh:
                    self.templates[f] = fh.read()

    # ── UI construction ──────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=SIDEBAR, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=ACCENT, width=5).pack(side="left", fill="y")
        tk.Label(hdr, text="  📧  EMAIL AUTOMATION",
                 bg=SIDEBAR, fg=TEXT, font=FONT_TITLE).pack(side="left", padx=(10, 0))
        tk.Label(hdr, text="smtplib + pandas",
                 bg=SIDEBAR, fg=SUBTEXT, font=FONT_SMALL).pack(side="left", padx=12)

        # ── Tabbed layout ─────────────────────────────────────────
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=16, pady=(14, 0))

        self.tab_smtp = self._make_smtp_tab(nb)
        self.tab_contacts = self._make_contacts_tab(nb)
        self.tab_compose = self._make_compose_tab(nb)

        # ── Log area ──────────────────────────────────────────────
        tk.Label(self.root, text="📋 Activity Log", bg=BG, fg=SUBTEXT,
                 font=FONT_BOLD).pack(anchor="w", padx=20, pady=(12, 4))
        log_wrap = tk.Frame(self.root, bg=CARD, highlightbackground=BORDER,
                            highlightthickness=1)
        log_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log_text = tk.Text(log_wrap, bg=CARD2, fg=TEXT, font=("Consolas", 9),
                                wrap="word", relief="flat", padx=12, pady=10,
                                state="disabled", height=7)
        self.log_text.pack(fill="both", expand=True)

    # ── Tab: SMTP ────────────────────────────────────────────────
    def _make_smtp_tab(self, nb):
        tab = tk.Frame(nb, bg=CARD)
        nb.add(tab, text="⚙ SMTP Settings")
        pad = tk.Frame(tab, bg=CARD)
        pad.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(pad, text="SMTP Server Settings", bg=CARD, fg=TEXT,
                 font=FONT_BOLD).pack(anchor="w", pady=(0, 10))

        # Server
        server_row = tk.Frame(pad, bg=CARD)
        server_row.pack(fill="x", pady=4)
        tk.Label(server_row, text="SMTP Server", width=16, bg=CARD, fg=SUBTEXT,
                 font=FONT).pack(side="left")
        self.server_var = tk.StringVar(value="smtp.gmail.com")
        tk.Entry(server_row, textvariable=self.server_var, bg=CARD2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(side="left", fill="x", expand=True, ipady=6)

        # Port + security
        port_row = tk.Frame(pad, bg=CARD)
        port_row.pack(fill="x", pady=4)
        tk.Label(port_row, text="Port", width=16, bg=CARD, fg=SUBTEXT,
                 font=FONT).pack(side="left")
        self.port_var = tk.StringVar(value="587")
        tk.Entry(port_row, textvariable=self.port_var, width=10, bg=CARD2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(side="left", ipady=6, ipadx=4)
        self.ssl_var = tk.BooleanVar(value=False)
        tk.Checkbutton(port_row, text="Use SSL/STARTTLS", variable=self.ssl_var,
                       bg=CARD, fg=TEXT, selectcolor=CARD2, font=FONT).pack(side="left", padx=16)

        # Sender email
        email_row = tk.Frame(pad, bg=CARD)
        email_row.pack(fill="x", pady=4)
        tk.Label(email_row, text="Sender Email", width=16, bg=CARD, fg=SUBTEXT,
                 font=FONT).pack(side="left")
        self.sender_var = tk.StringVar()
        tk.Entry(email_row, textvariable=self.sender_var, bg=CARD2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(side="left", fill="x", expand=True, ipady=6)

        # App password
        pass_row = tk.Frame(pad, bg=CARD)
        pass_row.pack(fill="x", pady=4)
        tk.Label(pass_row, text="App Password", width=16, bg=CARD, fg=SUBTEXT,
                 font=FONT).pack(side="left")
        self.pass_var = tk.StringVar()
        tk.Entry(pass_row, textvariable=self.pass_var, show="•", bg=CARD2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(side="left", fill="x", expand=True, ipady=6)

        tk.Label(pad, text="💡 Tip: For Gmail use an App Password (Settings → Security → 2-Step Verification → App passwords).",
                 bg=CARD, fg=SUBTEXT, font=FONT_SMALL).pack(anchor="w", pady=(14, 0))

        # Test connection
        tk.Button(pad, text="🔌 Test SMTP Connection", command=self._test_smtp,
                  bg=ACCENT, fg=WHITE, relief="flat", font=FONT_BOLD,
                  activebackground=BLUE, cursor="hand2", pady=8).pack(pady=(18, 0))
        return tab

    # ── Tab: Contacts ────────────────────────────────────────────
    def _make_contacts_tab(self, nb):
        tab = tk.Frame(nb, bg=CARD)
        nb.add(tab, text="👥 Contacts (Excel)")
        pad = tk.Frame(tab, bg=CARD, padx=16, pady=16)
        pad.pack(fill="both", expand=True)

        # Top row
        top = tk.Frame(pad, bg=CARD)
        top.pack(fill="x", pady=(0, 10))
        tk.Button(top, text="📂 Load Excel/CSV", command=self._load_contacts,
                  bg=ACCENT, fg=WHITE, relief="flat", font=FONT_BOLD,
                  activebackground=BLUE, cursor="hand2", padx=14, pady=6).pack(side="left")
        self.file_label = tk.Label(top, text="No file loaded", bg=CARD, fg=SUBTEXT,
                                   font=FONT).pack(side="left", padx=12)
        self.count_label = tk.Label(top, text="", bg=CARD, fg=GREEN, font=FONT_BOLD)
        self.count_label.pack(side="right")

        # Columns hint
        tk.Label(pad, text="Columns recognised as placeholders:  {{Name}}  {{Email}}  {{Company}}  {{Code}}  …",
                 bg=CARD, fg=SUBTEXT, font=FONT_SMALL).pack(anchor="w", pady=(0, 8))

        # Treeview
        frame = tk.Frame(pad, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill="both", expand=True)
        cols = ("Name", "Email", "Company", "Code")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="w", width=180)
        vs = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Bottom
        bottom = tk.Frame(pad, bg=CARD)
        bottom.pack(fill="x", pady=(10, 0))
        tk.Label(bottom, text="Message contains placeholders — they will be filled per contact.",
                 bg=CARD, fg=SUBTEXT, font=FONT_SMALL).pack(side="left")
        return tab

    # ── Tab: Compose ─────────────────────────────────────────────
    def _make_compose_tab(self, nb):
        tab = tk.Frame(nb, bg=CARD)
        nb.add(tab, text="✉ Compose & Send")
        pad = tk.Frame(tab, bg=CARD, padx=16, pady=16)
        pad.pack(fill="both", expand=True)

        # Template selector
        temp_row = tk.Frame(pad, bg=CARD)
        temp_row.pack(fill="x", pady=(0, 8))
        tk.Label(temp_row, text="Template:", bg=CARD, fg=SUBTEXT,
                 font=FONT_BOLD).pack(side="left")
        self.template_var = tk.StringVar()
        self.template_menu = ttk.Combobox(temp_row, textvariable=self.template_var,
                                          values=list(self.templates.keys()),
                                          state="readonly", width=30)
        self.template_menu.pack(side="left", padx=8)
        self.template_menu.bind("<<ComboboxSelected>>", self._on_template_selected)
        tk.Button(temp_row, text="↻ Refresh", command=self._refresh_templates,
                  bg=CARD2, fg=TEXT, relief="flat", font=FONT_BOLD, cursor="hand2").pack(side="left", padx=8)
        tk.Button(temp_row, text="📄 Load .txt", command=self._load_template_file,
                  bg=CARD2, fg=TEXT, relief="flat", font=FONT_BOLD, cursor="hand2").pack(side="left")

        # Subject
        subj_row = tk.Frame(pad, bg=CARD)
        subj_row.pack(fill="x", pady=(4, 8))
        tk.Label(subj_row, text="Subject:", bg=CARD, fg=SUBTEXT, font=FONT_BOLD).pack(side="left")
        self.subject_var = tk.StringVar()
        tk.Entry(subj_row, textvariable=self.subject_var, bg=CARD2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(side="left", fill="x", expand=True, padx=8, ipady=6)

        # Body
        tk.Label(pad, text="Body (use {{placeholder}} tokens):", bg=CARD, fg=SUBTEXT,
                 font=FONT_BOLD).pack(anchor="w", pady=(4, 6))
        body_frame = tk.Frame(pad, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        body_frame.pack(fill="both", expand=True)
        self.body_text = tk.Text(body_frame, bg=CARD2, fg=TEXT, font=("Consolas", 10),
                                 wrap="word", relief="flat", padx=10, pady=8,
                                 insertbackground=TEXT, undo=True)
        self.body_text.pack(fill="both", expand=True)

        # Attachments
        att = tk.Frame(pad, bg=CARD)
        att.pack(fill="x", pady=(10, 0))
        tk.Button(att, text="📎 Add Attachments", command=self._add_attachments,
                  bg=ACCENT, fg=WHITE, relief="flat", font=FONT_BOLD, cursor="hand2",
                  padx=12, pady=6).pack(side="left")
        self.att_label = tk.Label(att, text="No attachments", bg=CARD, fg=SUBTEXT, font=FONT)
        self.att_label.pack(side="left", padx=12)

        # Send
        send_row = tk.Frame(pad, bg=CARD)
        send_row.pack(fill="x", pady=(12, 0))
        self.send_btn = tk.Button(send_row, text="🚀 Send Bulk Email", command=self._send_bulk,
                                  bg=GREEN, fg="#0b1220", relief="flat", font=("Segoe UI", 12, "bold"),
                                  activebackground=ACCENT, cursor="hand2", pady=10)
        self.send_btn.pack(fill="x")
        return tab

    # ── Logging & helpers ────────────────────────────────────────
    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    # ── SMTP ─────────────────────────────────────────────────────
    def _smtp_connect(self):
        server = self.server_var.get().strip()
        port = int(self.port_var.get().strip())
        sender = self.sender_var.get().strip()
        password = self.pass_var.get()
        if not (server and sender and password):
            raise ValueError("Please fill in SMTP server, port, sender email and password.")
        if self.ssl_var.get():
            context = ssl.create_default_context()
            smtp = smtplib.SMTP_SSL(server, port, context=context, timeout=30)
        else:
            smtp = smtplib.SMTP(server, port, timeout=30)
            smtp.ehlo()
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
        smtp.login(sender, password)
        return smtp

    def _test_smtp(self):
        try:
            smtp = self._smtp_connect()
            smtp.quit()
            self._log("✅ SMTP connection successful.")
            messagebox.showinfo("Success", "SMTP connection successful!")
        except Exception as e:
            self._log(f"❌ SMTP test failed: {e}")
            messagebox.showerror("SMTP Error", str(e))

    # ── Contacts ─────────────────────────────────────────────────
    def _load_contacts(self):
        path = filedialog.askopenfilename(
            title="Select Contacts File",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self.contacts = pd.read_csv(path)
            else:
                self.contacts = pd.read_excel(path)
            self.file_label.config(text=os.path.basename(path))
            self.file_label.configure(fg=TEXT)
            # fill tree
            self.tree.delete(*self.tree.get_children())
            cols = list(self.contacts.columns)
            for i, c in enumerate(cols):
                self.tree.heading(i if i < len(self.tree["columns"]) else 0, text=c)
            for _, row in self.contacts.head(200).iterrows():
                self.tree.insert("", "end", values=[str(row[c]) for c in cols[:4]])
            self.count_label.config(text=f"{len(self.contacts)} contacts loaded")
            self._log(f"✅ Loaded {len(self.contacts)} contacts from {os.path.basename(path)}")
        except Exception as e:
            self._log(f"❌ Failed to load contacts: {e}")
            messagebox.showerror("Load Error", str(e))

    # ── Templates ────────────────────────────────────────────────
    def _refresh_templates(self):
        self._load_templates()
        self.template_menu["values"] = list(self.templates.keys())
        self._log("Templates refreshed.")

    def _on_template_selected(self, event=None):
        name = self.template_var.get()
        if name in self.templates:
            content = self.templates[name]
            subject = ""
            body = content
            if content.startswith("Subject:"):
                first, _, rest = content.partition("\n")
                subject = first.replace("Subject:", "").strip()
                body = rest.lstrip("\n")
            self.subject_var.set(subject)
            self.body_text.delete("1.0", "end")
            self.body_text.insert("1.0", body)

    def _load_template_file(self):
        path = filedialog.askopenfilename(title="Select Template",
                                          filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as fh:
            content = fh.read()
        self.templates[os.path.basename(path)] = content
        self.template_menu["values"] = list(self.templates.keys())
        self.template_var.set(os.path.basename(path))
        self._on_template_selected()
        self._log(f"Loaded template: {os.path.basename(path)}")

    # ── Attachments ──────────────────────────────────────────────
    def _add_attachments(self):
        paths = filedialog.askopenfilenames(title="Select Attachments")
        if paths:
            self.attachments.extend(paths)
            names = [os.path.basename(p) for p in self.attachments]
            self.att_label.config(text=f"{len(self.attachments)} file(s): " + ", ".join(names[:3]))
            self._log(f"📎 Added {len(paths)} attachment(s). Total: {len(self.attachments)}")

    # ── Build message ────────────────────────────────────────────
    def _build_message(self, row):
        subject = render_template(self.subject_var.get().strip(), row) or "(no subject)"
        body = render_template(self.body_text.get("1.0", "end").strip(), row)
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = formataddr((self.sender_var.get(), self.sender_var.get()))
        msg["To"] = row["Email"]
        msg.set_content(body)
        for path in self.attachments:
            with open(path, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="octet-stream",
                                   filename=os.path.basename(path))
        return msg

    # ── Send bulk ────────────────────────────────────────────────
    def _send_bulk(self):
        if self.contacts.empty:
            return messagebox.showwarning("No Contacts", "Please load a contacts file first.")
        if "Email" not in self.contacts.columns:
            return messagebox.showwarning("No Email Column",
                                          "Contacts file must have an 'Email' column.")
        # disable button, run in thread
        self._send_button_state(False)
        threading.Thread(target=self._send_worker, daemon=True).start()

    def _send_button_state(self, enabled):
        self.send_btn.config(state="normal" if enabled else "disabled",
                             text="🚀 Send Bulk Email" if enabled else "⏳ Sending…")
        if enabled:
            self._log("Ready.")

    def _send_worker(self):
        try:
            smtp = self._smtp_connect()
        except Exception as e:
            self._log(f"❌ Could not connect: {e}")
            messagebox.showerror("SMTP Error", str(e))
            self._send_button_state(True)
            return

        sent, failed = 0, 0
        total = len(self.contacts)
        self._log(f"🚀 Sending to {total} recipient(s)…")
        for idx, row in self.contacts.iterrows():
            email = row.get("Email")
            if not email or pd.isna(email):
                self._log(f"  ⏭ Skipped row {idx} (no email).")
                continue
            try:
                msg = self._build_message(row)
                smtp.send_message(msg)
                sent += 1
                self._log(f"  ✅ {email}")
            except Exception as e:
                failed += 1
                self._log(f"  ❌ {email} → {e}")
        smtp.quit()
        self._log("")
        self._log(f"--- Done! Sent: {sent}  |  Failed: {failed}  |  Total: {total} ---")
        self._send_button_state(True)
        messagebox.showinfo("Complete", f"Email campaign finished!\n\nSent: {sent}\nFailed: {failed}")


# ── Entry point ──────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    EmailAutomation(root)
    root.mainloop()
