# -*- coding: utf-8 -*-
import re
from sidebar.core.compat import tk, messagebox
from sidebar.core.config import VERSION

# GitHub Releases page for InboxBar
DOWNLOAD_URL = "https://github.com/Ralph180976/outlook_sidebar/releases/latest"


class ShareDialog(tk.Toplevel):
    """Dialog to share InboxBar via email with a download link."""

    def __init__(self, parent, outlook_client):
        tk.Toplevel.__init__(self, parent)
        self.outlook_client = outlook_client
        self.title("Share InboxBar")
        self.geometry("370x195")
        self.resizable(False, False)

        # Make Modal
        self.transient(parent)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()

        # Center on parent
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 185
            py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 97
            self.geometry("+{}+{}".format(px, py))
        except:
            pass

        # Dark Theme Colors
        self.colors = {
            "bg_root": "#202020",
            "bg_card": "#2D2D30",
            "bg_hover": "#3A3A3D",
            "accent": "#60CDFF",
            "fg_text": "#FFFFFF",
            "fg_dim": "#AAAAAA",
        }
        self.configure(
            bg=self.colors["bg_root"],
            highlightthickness=1,
            highlightbackground=self.colors["accent"],
        )

        # Content
        content = tk.Frame(self, bg=self.colors["bg_root"], padx=20, pady=15)
        content.pack(fill="both", expand=True)

        tk.Label(
            content,
            text="Share InboxBar with a colleague",
            fg=self.colors["fg_text"],
            bg=self.colors["bg_root"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))

        tk.Label(
            content,
            text="Enter their email address to send a download link:",
            fg=self.colors["fg_dim"],
            bg=self.colors["bg_root"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        # Email Entry
        self.email_var = tk.StringVar()
        self.email_entry = tk.Entry(
            content,
            textvariable=self.email_var,
            bg=self.colors["bg_card"],
            fg=self.colors["fg_text"],
            insertbackground="white",
            font=("Segoe UI", 10),
            relief="flat",
            bd=0,
        )
        self.email_entry.pack(fill="x", ipady=6, pady=(0, 15))
        self.email_entry.focus_set()
        self.email_entry.bind("<Return>", lambda e: self._on_send())

        # Buttons
        btn_frame = tk.Frame(content, bg=self.colors["bg_root"])
        btn_frame.pack(fill="x")

        btn_cancel = tk.Label(
            btn_frame,
            text="Cancel",
            fg="#AAAAAA",
            bg=self.colors["bg_root"],
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=15,
            pady=5,
        )
        btn_cancel.pack(side="right", padx=5)
        btn_cancel.bind("<Button-1>", lambda e: self.destroy())

        btn_send = tk.Label(
            btn_frame,
            text="  Send  ",
            fg="white",
            bg=self.colors["accent"],
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=15,
            pady=5,
        )
        btn_send.pack(side="right", padx=5)
        btn_send.bind("<Button-1>", lambda e: self._on_send())

        # Hover effects
        btn_send.bind("<Enter>", lambda e: btn_send.config(bg="#40b0ff"))
        btn_send.bind("<Leave>", lambda e: btn_send.config(bg=self.colors["accent"]))
        btn_cancel.bind("<Enter>", lambda e: btn_cancel.config(fg="white"))
        btn_cancel.bind("<Leave>", lambda e: btn_cancel.config(fg="#AAAAAA"))

        # Toast label (hidden initially)
        self._toast = tk.Label(
            self,
            text="",
            fg="white",
            bg="#2E7D32",
            font=("Segoe UI", 10),
            anchor="center",
            pady=6,
        )

    def _show_toast(self, message, success=True):
        """Show an auto-dismissing toast notification, then close the dialog."""
        color = "#2E7D32" if success else "#C62828"
        self._toast.config(text=message, bg=color)
        self._toast.place(x=0, y=0, relwidth=1.0)
        self._toast.lift()
        # Auto-dismiss and close after 2 seconds
        self.after(2000, self._dismiss)
    
    def _dismiss(self):
        """Close the dialog."""
        try:
            self.destroy()
        except:
            pass

    def _on_send(self):
        email = self.email_var.get().strip()

        # Validate email
        if not email:
            self._show_toast("Please enter an email address", success=False)
            return
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            self._show_toast("Invalid email address", success=False)
            return

        # Build email with download link
        subject = "Try InboxBar - Outlook Sidebar ({})".format(VERSION)
        body = (
            "Hi,\n\n"
            "I'd like to share InboxBar with you - a streamlined sidebar for Outlook "
            "that keeps you updated without the clutter of the full Outlook window.\n\n"
            "Download the latest version here:\n"
            "{}\n\n"
            "To install:\n"
            "1. Download InboxBar_Setup and run the installer\n"
            "2. InboxBar will appear in your Start Menu and optionally on your Desktop\n\n"
            "To uninstall, use Add/Remove Programs in Windows Settings.\n\n"
            "Requirements: Windows 10/11 + Microsoft Outlook (Classic)\n\n"
            "Enjoy!"
        ).format(DOWNLOAD_URL)

        # Send email
        try:
            mail = self.outlook_client.outlook.CreateItem(0)  # olMailItem
            mail.To = email
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            self._show_toast("Sent to {}".format(email), success=True)
        except Exception as e:
            print("Share email error: {}".format(e))
            self._show_toast("Failed to send. Check Outlook.", success=False)
