# -*- coding: utf-8 -*-
from sidebar.core.compat import tk
from sidebar.core.config import VERSION

# Direct download link for the latest InboxBar installer
# (will be replaced with Microsoft Store link later)
DOWNLOAD_URL = "https://github.com/Ralph180976/outlook_sidebar/releases/latest/download/InboxBar_Setup.exe"


class ShareDialog(tk.Toplevel):
    """Dialog to share InboxBar by copying a download link."""

    def __init__(self, parent, outlook_client=None):
        tk.Toplevel.__init__(self, parent)
        self.title("Share InboxBar")
        self.geometry("380x220")
        self.resizable(False, False)

        # Make Modal
        self.transient(parent)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()

        # Center on parent
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 190
            py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 110
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
            "success": "#2E7D32",
        }
        self.configure(
            bg=self.colors["bg_root"],
            highlightthickness=1,
            highlightbackground=self.colors["accent"],
        )

        # --- Content ---
        content = tk.Frame(self, bg=self.colors["bg_root"], padx=20, pady=15)
        content.pack(fill="both", expand=True)

        tk.Label(
            content,
            text="Share InboxBar",
            fg=self.colors["fg_text"],
            bg=self.colors["bg_root"],
            font=("Segoe UI", 12, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))

        tk.Label(
            content,
            text="Copy the download link below and paste it\ninto an email or message:",
            fg=self.colors["fg_dim"],
            bg=self.colors["bg_root"],
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(0, 12))

        # Link display (read-only entry)
        link_frame = tk.Frame(content, bg=self.colors["accent"], padx=1, pady=1)
        link_frame.pack(fill="x", pady=(0, 15))

        self.link_entry = tk.Entry(
            link_frame,
            bg=self.colors["bg_card"],
            fg=self.colors["accent"],
            insertbackground=self.colors["accent"],
            font=("Segoe UI", 9),
            relief="flat",
            bd=0,
            readonlybackground=self.colors["bg_card"],
        )
        self.link_entry.pack(fill="x", ipady=6, padx=1, pady=1)
        self.link_entry.insert(0, DOWNLOAD_URL)
        self.link_entry.config(state="readonly")

        # Buttons
        btn_frame = tk.Frame(content, bg=self.colors["bg_root"])
        btn_frame.pack(fill="x")

        btn_close = tk.Label(
            btn_frame,
            text="Close",
            fg="#AAAAAA",
            bg=self.colors["bg_root"],
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=15,
            pady=5,
        )
        btn_close.pack(side="right", padx=5)
        btn_close.bind("<Button-1>", lambda e: self.destroy())

        self.btn_copy = tk.Label(
            btn_frame,
            text="  \U0001F4CB  Copy Download Link  ",
            fg="white",
            bg=self.colors["accent"],
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=15,
            pady=5,
        )
        self.btn_copy.pack(side="right", padx=5)
        self.btn_copy.bind("<Button-1>", lambda e: self._copy_link())

        # Hover effects
        self.btn_copy.bind("<Enter>", lambda e: self.btn_copy.config(bg="#40b0ff"))
        self.btn_copy.bind("<Leave>", lambda e: self._reset_copy_btn())
        btn_close.bind("<Enter>", lambda e: btn_close.config(fg="white"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(fg="#AAAAAA"))

        # Track copied state for hover reset
        self._copied = False

        # Toast label (hidden initially)
        self._toast = tk.Label(
            self,
            text="",
            fg="white",
            bg=self.colors["success"],
            font=("Segoe UI", 10),
            anchor="center",
            pady=6,
        )

    def _reset_copy_btn(self):
        """Reset button color on mouse leave."""
        if self._copied:
            self.btn_copy.config(bg=self.colors["success"])
        else:
            self.btn_copy.config(bg=self.colors["accent"])

    def _copy_link(self):
        """Copy the download URL to the clipboard."""
        try:
            self.clipboard_clear()
            self.clipboard_append(DOWNLOAD_URL)
            self.update()  # Required for clipboard to persist

            # Visual feedback - change button text and color
            self._copied = True
            self.btn_copy.config(
                text="  \u2713  Copied!  ",
                bg=self.colors["success"],
            )
            # Reset after 2 seconds
            self.after(2000, self._reset_copy_state)
        except Exception as e:
            print("Copy error: {}".format(e))
            self._show_toast("Failed to copy to clipboard", success=False)

    def _reset_copy_state(self):
        """Reset the copy button to its original state."""
        self._copied = False
        try:
            self.btn_copy.config(
                text="  \U0001F4CB  Copy Download Link  ",
                bg=self.colors["accent"],
            )
        except:
            pass

    def _show_toast(self, message, success=True):
        """Show an auto-dismissing toast notification."""
        color = self.colors["success"] if success else "#C62828"
        self._toast.config(text=message, bg=color)
        self._toast.place(x=0, y=0, relwidth=1.0)
        self._toast.lift()
        self.after(2000, self._dismiss_toast)

    def _dismiss_toast(self):
        """Hide the toast notification."""
        try:
            self._toast.place_forget()
        except:
            pass
