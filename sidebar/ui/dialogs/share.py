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
        
        self._search_job = None  # Debounce timer for autocomplete
        self._suggestions = []   # Current suggestion results

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

        # Email Entry (wrapped in a frame for autocomplete positioning)
        self.entry_frame = tk.Frame(content, bg=self.colors["bg_root"])
        self.entry_frame.pack(fill="x", pady=(0, 15))
        
        self.email_var = tk.StringVar()
        self.email_entry = tk.Entry(
            self.entry_frame,
            textvariable=self.email_var,
            bg=self.colors["bg_card"],
            fg=self.colors["fg_text"],
            insertbackground="white",
            font=("Segoe UI", 10),
            relief="flat",
            bd=0,
        )
        self.email_entry.pack(fill="x", ipady=6)
        self.email_entry.focus_set()

        # Autocomplete Listbox (hidden initially, overlays below entry)
        self.suggest_frame = tk.Frame(
            self, bg=self.colors["accent"], bd=0,
        )
        self.suggest_list = tk.Listbox(
            self.suggest_frame,
            bg=self.colors["bg_card"],
            fg=self.colors["fg_text"],
            selectbackground=self.colors["accent"],
            selectforeground="#000000",
            font=("Segoe UI", 9),
            relief="flat",
            bd=0,
            highlightthickness=0,
            activestyle="none",
            cursor="hand2",
        )
        self.suggest_list.pack(fill="both", expand=True, padx=1, pady=1)
        
        # Bindings
        self.email_entry.bind("<KeyRelease>", self._on_key_release)
        self.email_entry.bind("<Return>", lambda e: self._on_send())
        self.email_entry.bind("<Down>", self._focus_suggestions)
        self.email_entry.bind("<Escape>", lambda e: self._hide_suggestions())
        
        self.suggest_list.bind("<Return>", self._select_suggestion)
        self.suggest_list.bind("<ButtonRelease-1>", self._select_suggestion)
        self.suggest_list.bind("<Escape>", self._return_to_entry)
        self.suggest_list.bind("<Up>", self._suggest_nav_up)

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

    # --- Autocomplete ---
    
    def _on_key_release(self, event):
        """Debounced search on keystroke."""
        # Ignore navigation keys
        if event.keysym in ("Up", "Down", "Left", "Right", "Escape", "Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        
        # Cancel pending search
        if self._search_job:
            self.after_cancel(self._search_job)
        
        query = self.email_var.get().strip()
        if len(query) < 2:
            self._hide_suggestions()
            return
        
        # Debounce: wait 300ms after last keystroke before searching
        self._search_job = self.after(300, lambda: self._do_search(query))
    
    def _do_search(self, query):
        """Execute the contact search and show results."""
        try:
            results = self.outlook_client.search_contacts(query)
            self._suggestions = results
            
            if not results:
                self._hide_suggestions()
                return
            
            self.suggest_list.delete(0, tk.END)
            for r in results:
                if r["name"] and r["name"] != r["email"]:
                    display = "{} <{}>".format(r["name"], r["email"])
                else:
                    display = r["email"]
                self.suggest_list.insert(tk.END, display)
            
            self._show_suggestions(len(results))
        except Exception as e:
            print("Autocomplete error: {}".format(e))
            self._hide_suggestions()
    
    def _show_suggestions(self, count):
        """Position and show the suggestion dropdown."""
        # Position directly below the entry field
        entry_x = self.email_entry.winfo_rootx() - self.winfo_rootx()
        entry_y = self.email_entry.winfo_rooty() - self.winfo_rooty() + self.email_entry.winfo_height()
        entry_w = self.email_entry.winfo_width()
        
        item_h = 22
        list_h = min(count, 6) * item_h + 2  # Max 6 visible items
        self.suggest_list.config(height=min(count, 6))
        
        self.suggest_frame.place(
            x=entry_x, y=entry_y,
            width=entry_w, height=list_h
        )
        self.suggest_frame.lift()
    
    def _hide_suggestions(self):
        """Hide the suggestion dropdown."""
        self.suggest_frame.place_forget()
        self._suggestions = []
    
    def _focus_suggestions(self, event=None):
        """Move focus to the suggestion list."""
        if self._suggestions:
            self.suggest_list.focus_set()
            self.suggest_list.selection_clear(0, tk.END)
            self.suggest_list.selection_set(0)
            self.suggest_list.activate(0)
        return "break"
    
    def _suggest_nav_up(self, event=None):
        """When pressing Up on first item, return to entry."""
        sel = self.suggest_list.curselection()
        if sel and sel[0] == 0:
            self._return_to_entry()
            return "break"
    
    def _return_to_entry(self, event=None):
        """Return focus to the email entry."""
        self.email_entry.focus_set()
        self.email_entry.icursor(tk.END)
        return "break"
    
    def _select_suggestion(self, event=None):
        """Select a suggestion and fill the email entry."""
        sel = self.suggest_list.curselection()
        if not sel:
            return
        
        idx = sel[0]
        if idx < len(self._suggestions):
            email = self._suggestions[idx]["email"]
            self.email_var.set(email)
            self.email_entry.icursor(tk.END)
        
        self._hide_suggestions()
        self.email_entry.focus_set()

    # --- Send ---

    def _on_send(self):
        self._hide_suggestions()
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
            "1. Download the zip file and extract it\n"
            "2. Run the InboxBar Setup installer\n"
            "3. InboxBar will appear in your Start Menu and optionally on your Desktop\n\n"
            "To uninstall, use Add/Remove Programs in Windows Settings.\n\n"
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
