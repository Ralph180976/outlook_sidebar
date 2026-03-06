# -*- coding: utf-8 -*-
from sidebar.core.compat import tk
from sidebar.core.config import VERSION

# Direct download link for the latest InboxBar installer
# (will be replaced with Microsoft Store link later)
DOWNLOAD_URL = "https://github.com/Ralph180976/outlook_sidebar/releases/latest/download/InboxBar_Setup.exe"

LINK_TEXT = "InboxBar Download"

SHARE_MESSAGE = (
    "Hi,\n"
    "\n"
    "I've been using InboxBar - a lightweight sidebar for Outlook that shows "
    "your emails, calendar and tasks at a glance.\n"
    "\n"
    "You can download it here: {url}\n"
).format(url=DOWNLOAD_URL)


def _copy_html_link(widget):
    """Copy a clickable hyperlink to the Windows clipboard using CF_HTML.
    
    When pasted into Outlook, Teams, etc. it appears as a clickable
    'InboxBar Download' link rather than a raw URL.
    Falls back to plain text if win32clipboard is not available.
    """
    html_body = '<a href="{url}">{text}</a>'.format(url=DOWNLOAD_URL, text=LINK_TEXT)
    
    try:
        import win32clipboard
        
        # Build CF_HTML payload (requires specific header)
        html_template = (
            "Version:0.9\r\n"
            "StartHTML:{start_html:08d}\r\n"
            "EndHTML:{end_html:08d}\r\n"
            "StartFragment:{start_frag:08d}\r\n"
            "EndFragment:{end_frag:08d}\r\n"
            "<html><body>\r\n"
            "<!--StartFragment-->{fragment}<!--EndFragment-->\r\n"
            "</body></html>"
        )
        
        # Calculate offsets - need to do two passes since header length depends on content
        dummy = html_template.format(
            start_html=0, end_html=0, start_frag=0, end_frag=0,
            fragment=html_body
        )
        # Now calculate real positions
        header_end = dummy.index("<html>")
        start_html = header_end
        start_frag = dummy.index("<!--StartFragment-->") + len("<!--StartFragment-->")
        end_frag = dummy.index("<!--EndFragment-->")
        end_html = len(dummy)
        
        cf_html = html_template.format(
            start_html=start_html,
            end_html=end_html,
            start_frag=start_frag,
            end_frag=end_frag,
            fragment=html_body
        )
        
        CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        # Set both HTML and plain text so it pastes nicely everywhere
        win32clipboard.SetClipboardData(CF_HTML, cf_html.encode("utf-8"))
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, DOWNLOAD_URL)
        win32clipboard.CloseClipboard()
        return True
        
    except Exception as e:
        print("HTML clipboard error, falling back to plain text: {}".format(e))
        # Fallback: plain text via tkinter
        widget.clipboard_clear()
        widget.clipboard_append(DOWNLOAD_URL)
        widget.update()
        return True


def _copy_html_message(widget):
    """Copy a message with a clickable hyperlink to the Windows clipboard."""
    html_link = '<a href="{url}">{text}</a>'.format(url=DOWNLOAD_URL, text=LINK_TEXT)
    
    html_body = (
        "Hi,<br><br>"
        "I've been using InboxBar - a lightweight sidebar for Outlook that shows "
        "your emails, calendar and tasks at a glance.<br><br>"
        "You can download it here: {link}<br>"
    ).format(link=html_link)
    
    try:
        import win32clipboard
        
        html_template = (
            "Version:0.9\r\n"
            "StartHTML:{start_html:08d}\r\n"
            "EndHTML:{end_html:08d}\r\n"
            "StartFragment:{start_frag:08d}\r\n"
            "EndFragment:{end_frag:08d}\r\n"
            "<html><body>\r\n"
            "<!--StartFragment-->{fragment}<!--EndFragment-->\r\n"
            "</body></html>"
        )
        
        dummy = html_template.format(
            start_html=0, end_html=0, start_frag=0, end_frag=0,
            fragment=html_body
        )
        header_end = dummy.index("<html>")
        start_html = header_end
        start_frag = dummy.index("<!--StartFragment-->") + len("<!--StartFragment-->")
        end_frag = dummy.index("<!--EndFragment-->")
        end_html = len(dummy)
        
        cf_html = html_template.format(
            start_html=start_html,
            end_html=end_html,
            start_frag=start_frag,
            end_frag=end_frag,
            fragment=html_body
        )
        
        CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(CF_HTML, cf_html.encode("utf-8"))
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, SHARE_MESSAGE.strip())
        win32clipboard.CloseClipboard()
        return True
        
    except Exception as e:
        print("HTML clipboard error, falling back to plain text: {}".format(e))
        widget.clipboard_clear()
        widget.clipboard_append(SHARE_MESSAGE.strip())
        widget.update()
        return True


class ShareDialog(tk.Toplevel):
    """Dialog to share InboxBar by copying a download link or message."""

    def __init__(self, parent, outlook_client=None):
        tk.Toplevel.__init__(self, parent)
        self.title("Share InboxBar")
        self.geometry("340x185")
        self.resizable(False, False)

        # Make Modal
        self.transient(parent)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()

        # Center on parent
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 170
            py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 92
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
            text="Share with a colleague via email or message:",
            fg=self.colors["fg_dim"],
            bg=self.colors["bg_root"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(0, 18))

        # --- Buttons ---
        btn_frame = tk.Frame(content, bg=self.colors["bg_root"])
        btn_frame.pack(fill="x")

        # Copy Link button
        self.btn_link = tk.Label(
            btn_frame,
            text="  Copy Link  ",
            fg="white",
            bg=self.colors["accent"],
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=12,
            pady=6,
        )
        self.btn_link.pack(side="left", padx=(0, 8))
        self.btn_link.bind("<Button-1>", lambda e: self._on_copy_link())
        self.btn_link.bind("<Enter>", lambda e: self._hover(self.btn_link, True))
        self.btn_link.bind("<Leave>", lambda e: self._hover(self.btn_link, False))

        # Copy Message button
        self.btn_msg = tk.Label(
            btn_frame,
            text="  Copy Message  ",
            fg="white",
            bg=self.colors["accent"],
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=12,
            pady=6,
        )
        self.btn_msg.pack(side="left", padx=(0, 8))
        self.btn_msg.bind("<Button-1>", lambda e: self._on_copy_message())
        self.btn_msg.bind("<Enter>", lambda e: self._hover(self.btn_msg, True))
        self.btn_msg.bind("<Leave>", lambda e: self._hover(self.btn_msg, False))

        # Close button
        btn_close = tk.Label(
            btn_frame,
            text="Close",
            fg="#AAAAAA",
            bg=self.colors["bg_root"],
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=10,
            pady=6,
        )
        btn_close.pack(side="right")
        btn_close.bind("<Button-1>", lambda e: self.destroy())
        btn_close.bind("<Enter>", lambda e: btn_close.config(fg="white"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(fg="#AAAAAA"))

        # Track which button is in "copied" state
        self._active_btn = None

    def _hover(self, btn, entering):
        """Handle hover effect, respecting the copied state."""
        if btn == self._active_btn:
            btn.config(bg=self.colors["success"])
        else:
            btn.config(bg="#40b0ff" if entering else self.colors["accent"])

    def _on_copy_link(self):
        """Copy just the clickable 'InboxBar Download' link."""
        _copy_html_link(self)
        self._show_copied(self.btn_link)

    def _on_copy_message(self):
        """Copy a short message with the clickable link."""
        _copy_html_message(self)
        self._show_copied(self.btn_msg)

    def _show_copied(self, btn):
        """Flash the button green with a tick to confirm copy."""
        # Reset previous button if any
        if self._active_btn and self._active_btn != btn:
            self._reset_btn(self._active_btn)

        self._active_btn = btn
        original_text = btn.cget("text")
        btn.config(text="  \u2713  Copied!  ", bg=self.colors["success"])
        self.after(2000, lambda: self._reset_btn(btn, original_text))

    def _reset_btn(self, btn, text=None):
        """Reset a button to its normal state."""
        if self._active_btn == btn:
            self._active_btn = None
        try:
            if text:
                btn.config(text=text, bg=self.colors["accent"])
            else:
                btn.config(bg=self.colors["accent"])
        except:
            pass
