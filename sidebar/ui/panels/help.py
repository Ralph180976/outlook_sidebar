# -*- coding: utf-8 -*-
from sidebar.core.compat import tk, ttk, messagebox
from PIL import Image, ImageTk
import os

from sidebar.core.config import RESAMPLE_MODE
from sidebar.ui.widgets.base import ScrollableFrame, ToolTip
from sidebar.ui.panels.account_settings import AccountSelectionDialog

class HelpPanel(tk.Frame):
    """Inline help panel that extends from the sidebar."""
    def __init__(self, parent, main_window):
        self.main_window = main_window
        # Inherit colors from Main Window
        self.colors = main_window.colors
        
        tk.Frame.__init__(self, parent, bg=self.colors["bg_root"])
        
        # Frame styling
        self.config(bg=self.colors["bg_root"])
        self.configure(highlightbackground=self.colors["divider"], highlightthickness=1)
        
        # Fixed width matches Settings
        self.panel_width = 370
        self.config(width=self.panel_width)
        self.pack_propagate(False)
        
        # --- Header ---
        header = tk.Frame(self, bg=self.colors["bg_root"], height=40)
        header.pack(fill="x", side="top")
        
        lbl_title = tk.Label(header, text="Instructions", fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Segoe UI Variable Display", 12, "bold"))
        lbl_title.pack(side="left", padx=20, pady=10)
        
        title_underline = tk.Frame(self, bg=self.colors["accent"], height=2)
        title_underline.pack(fill="x", side="top")
        
        # Close Button
        btn_close = tk.Label(header, text="✕", fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Arial", 14, "bold"), cursor="hand2")
        btn_close.pack(side="right", padx=10)
        btn_close.bind("<Button-1>", lambda e: self.main_window.toggle_help_panel())

        # --- Scrollable Container ---
        self.scroll_frame = ScrollableFrame(self, bg=self.colors["bg_root"], auto_hide_scrollbar=False)
        self.scroll_frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        main_content = self.scroll_frame.scrollable_frame
        main_content.config(bg=self.colors["bg_root"])

        # --- Content Sections ---
        
        # Image cache to prevent GC
        self.help_images = []

        def create_text_section(parent, header, text, image_path=None):
            frame = tk.Frame(parent, bg=self.colors["bg_root"])
            frame.pack(fill="x", anchor="w", padx=15, pady=10)
            
            # Header
            tk.Label(frame, text=header, fg=self.colors["accent"], bg=self.colors["bg_root"], 
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            
            # Text
            tk.Label(frame, text=text, fg=self.colors["fg_dim"], bg=self.colors["bg_root"], 
                     font=("Segoe UI", 9), justify="left", wraplength=320).pack(anchor="w", pady=(5, 5))
            
            # Image Loading
            if image_path:
                 full_path = os.path.join("images", image_path)
                 if os.path.exists(full_path):
                     try:
                         # Load + Resize
                         pil_img = Image.open(full_path)
                         
                         # Calculate resize (Max width 320)
                         w_orig, h_orig = pil_img.size
                         max_w = 320
                         if w_orig > max_w:
                             ratio = float(max_w) / w_orig
                             new_h = int(h_orig * ratio)
                             pil_img = pil_img.resize((max_w, new_h), RESAMPLE_MODE)
                         
                         tk_img = ImageTk.PhotoImage(pil_img)
                         
                         # Cache it!
                         self.help_images.append(tk_img)
                         
                         img_lbl = tk.Label(frame, image=tk_img, bg=self.colors["bg_root"], bd=1, relief="solid")
                         img_lbl.pack(anchor="w", pady=5)
                         
                     except Exception as e:
                         print("Error loading help image {}: {}".format(full_path, e))
                         # Fallback to placeholder text
                         ph = tk.Canvas(frame, bg="#333333", height=100, highlightthickness=0)
                         ph.pack(fill="x", pady=5)
                         ph.create_text(160, 50, text="FAILED TO LOAD:\n{}".format(image_path), fill="#FF4444", font=("Segoe UI", 8), justify="center")
                 else:
                     # Placeholder for missing file
                     ph = tk.Canvas(frame, bg="#333333", height=100, highlightthickness=0)
                     ph.pack(fill="x", pady=5)
                     ph.create_text(160, 50, text="MISSING FILE:\n{}".format(image_path), fill="#666666", font=("Segoe UI", 8), justify="center")

        # 1. Introduction
        create_text_section(main_content, "Welcome to Outlook Sidebar", 
            "A streamlined, distraction-free interface for your Outlook emails, calendar, and tasks. It sits quietly on the edge of your screen, keeping you updated without the clutter of the full Outlook window.")

        # 2. The Header (Top Bar)
        create_text_section(main_content, "The Header",
            "The top bar gives you quick access to essential controls: Pin/Unpin the window, open Settings, Refresh data, or Share.",
            "Top Bar.png")

        # 3. Window Modes
        create_text_section(main_content, "Window Modes", 
            "Choose between 'Email Only' for a compact view, or 'Emails & Reminders' to see your upcoming schedule and tasks side-by-side.", 
            "Window Selection.png")

        # 4. General Settings
        create_text_section(main_content, "General Settings",
            "Customize the Look & Feel. Adjust the Font Family, Font Size, and how often the data refreshes.",
            "General Settings.png")

        # 5. Email List Settings
        create_text_section(main_content, "Email List Settings",
            "Control what you see. Toggle 'Read' emails, attachment icons, and even preview message content directly in the list.",
            "Email Settings.png")

        # 6. Quick Actions (Hover)
        create_text_section(main_content, "Quick Actions", 
            "Hover over any email to reveal quick actions like Reply, Delete, or Mark as Read.",
            "Hover Buttons.png")

        # 7. Customizing Buttons
        create_text_section(main_content, "Customizing Buttons", 
            "In Settings, assign different actions to the four quick-action slots. You can even set a button to 'Move To Folder' for one-click filing.",
            "Hover Button Settings.png")

        # 8. Reminders & Flags
        create_text_section(main_content, "Reminders & Flags", 
            "The Reminders pane shows flagged emails and Outlook Tasks. Use the filter options (Today, Tomorrow, Overdue) to focus on what's important now.",
            "Flagged_Reminders Window.png")

        create_text_section(main_content, "Reminder Settings",
            "Choose which reminders appear. You can filter by Follow-up Flags, Categories, or Importance.",
            "Reminder Settings.png")

        # 9. Quick Create
        create_text_section(main_content, "Quick Create", 
            "Use the icons at the bottom of the sidebar to instantly create new items. You can customize which actions appear here.",
            "Bottom Bar.png")

        create_text_section(main_content, "Quick Create Settings",
            "Select which creations you use most. \n\nNOTE: If only one option is selected (e.g., 'New Email'), the Quick Create button will immediately launch that action instead of showing a selection menu.",
            "Quick Create Settings.png")
            
    def open_account_selection(self):
        """Opens the account selection dialog."""
        accounts = self.main_window.outlook_client.get_accounts()
        if not accounts:
            messagebox.showerror("Error", "Could not fetch Outlook accounts.")
            return

        def on_save(new_settings):
            self.main_window.enabled_accounts = new_settings
            self.main_window.save_config()
            self.main_window.refresh_emails()
            self.main_window.refresh_reminders()
            
        AccountSelectionDialog(self.winfo_toplevel(), accounts, self.main_window.enabled_accounts, on_save, colors=self.colors)
