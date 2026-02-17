# -*- coding: utf-8 -*-
import re
from sidebar.core.compat import tk, messagebox
from sidebar.core.config import VERSION

# OneDrive direct download link for InboxBar installer
DOWNLOAD_URL = "https://1drv.ms/u/c/4aff08b047e14317/IQDcvbURHB01SKL1N_CG3i41AaMPi1rLeZAINjdmxVb9JW0?e=lAjsoq"


class ShareDialog(tk.Toplevel):
    """Dialog to share InboxBar via email with a download link."""

    def __init__(self, parent, outlook_client):
        tk.Toplevel.__init__(self, parent)
        self.outlook_client = outlook_client
        self.title("Share InboxBar")
        self.geometry("370x210")
        self.resizable(False, False)

        # Make Modal
        self.transient(parent)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()

        # Center on parent
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 185
            py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 105
            self.geometry("+{}+{}".format(px, py))
        except:
            pass

        # Dark Theme Colors
        self.colors = {
            "bg_root": "#202020",
            "bg_card": "#2D2D30",
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

        # Bind Enter key to send
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

    def _on_send(self):
        email = self.email_var.get().strip()

        # Validate email
        if not email:
            messagebox.showwarning("Missing Email", "Please enter an email address.")
            return
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            messagebox.showwarning(
                "Invalid Email", "Please enter a valid email address."
            )
            return

        # Build email with download link (no attachment)
        subject = "Try InboxBar - Outlook Sidebar ({})".format(VERSION)
        body = (
            "Hi,\n\n"
            "I'd like to share InboxBar with you - a streamlined sidebar for Outlook "
            "that keeps you updated without the clutter of the full Outlook window.\n\n"
            "Download the installer here:\n"
            "{}\n\n"
            "To install:\n"
            "1. Download and extract the zip file to a folder\n"
            "2. Double-click Setup.bat\n"
            "3. Launch InboxBar from your desktop\n\n"
            "Requirements: Windows 10/11 + Microsoft Outlook (Classic)\n\n"
            "Enjoy!"
        ).format(DOWNLOAD_URL)

        # Send email without attachment
        try:
            mail = self.outlook_client.outlook.CreateItem(0)  # olMailItem
            mail.To = email
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            success = True
        except Exception as e:
            print("Share email error: {}".format(e))
            success = False

        if success:
            messagebox.showinfo("Sent!", "InboxBar download link sent to {}".format(email))
            self.destroy()
        else:
            messagebox.showerror(
                "Error",
                "Failed to send email.\nPlease check Outlook is running.",
            )
