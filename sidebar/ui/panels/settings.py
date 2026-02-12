# -*- coding: utf-8 -*-
from sidebar.core.compat import tk, ttk, messagebox
import os
import glob
from PIL import Image, ImageTk

from sidebar.ui.widgets.base import ScrollableFrame, ToolTip
from sidebar.ui.panels.account_settings import FolderPickerWindow, FolderPickerFrame

class SettingsPanel(tk.Frame):
    """Inline settings panel that extends from the sidebar."""
    def __init__(self, parent, main_window, callback):
        tk.Frame.__init__(self, parent, bg="#202020")
        self.main_window = main_window
        self.callback = callback
        
        # Inherit colors from Main Window (Theme Aware)
        self.colors = main_window.colors
        
        # Configure ttk Theme
        style = ttk.Style(self)
        style.theme_use("clam")
        
        # TCombobox - Flat, Dynamic
        style.configure("TCombobox", 
            fieldbackground=self.colors["input_bg"], 
            background=self.colors["bg_card"], 
            foreground=self.colors["fg_text"],
            arrowcolor=self.colors["fg_text"],
            bordercolor=self.colors["bg_root"],
            darkcolor=self.colors["bg_root"],
            lightcolor=self.colors["bg_root"]
        )
        style.map("TCombobox", fieldbackground=[("readonly", self.colors["input_bg"])])
        
        # TEntry - Flat, Dynamic
        style.configure("TEntry", 
            fieldbackground=self.colors["input_bg"], 
            foreground=self.colors["fg_text"],
            bordercolor=self.colors["bg_root"],
            lightcolor=self.colors["bg_root"],
            darkcolor=self.colors["bg_root"]
        )
        
        # Frame styling
        self.config(bg=self.colors["bg_root"])
        self.configure(highlightbackground=self.colors["divider"], highlightthickness=1)
        
        # Fixed width for the settings panel
        self.panel_width = 370
        self.config(width=self.panel_width)
        self.pack_propagate(False)  # Prevent shrinking
        
        # --- Header ---
        header = tk.Frame(self, bg=self.colors["bg_root"], height=40)
        header.pack(fill="x", side="top")
        
        lbl_title = tk.Label(header, text="Settings", fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Segoe UI Variable Display", 12, "bold"))
        lbl_title.pack(side="left", padx=20, pady=10)
        
        # Theme Toggle
        theme_icon = "☀" if self.main_window.current_theme == "Dark" else "☾"
        btn_theme = tk.Label(header, text=theme_icon, fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Segoe UI Symbol", 14), cursor="hand2")
        btn_theme.pack(side="right", padx=10)
        btn_theme.bind("<Button-1>", lambda e: self.main_window.toggle_theme())
        ToolTip(btn_theme, "Toggle Light/Dark Mode")

        title_underline = tk.Frame(self, bg=self.colors["accent"], height=2)
        title_underline.pack(fill="x", side="top")

        # Configure larger font for dropdown lists
        self.option_add('*TCombobox*Listbox.font', ("Segoe UI", 10))

        # Create dedicated style for Font Size combobox (Dynamic)
        style.configure('FontSize.TCombobox',
            fieldbackground=self.colors["input_bg"],
            background=self.colors["bg_card"],
            foreground=self.colors["fg_text"],
            arrowcolor=self.colors["fg_text"],
            bordercolor=self.colors["divider"],
            lightcolor=self.colors["bg_root"],
            darkcolor=self.colors["bg_root"],
            selectbackground=self.colors["accent"],
            selectforeground='white'
        )

        
        # Map foreground color for readonly state (critical for visibility!)
        style.map('FontSize.TCombobox',
            fieldbackground=[('readonly', '#2d2d2d')],
            selectbackground=[('readonly', '#2d2d2d')],
            foreground=[('readonly', 'white')]  # Ensures white text in readonly mode
        )
        
        # Configure the listbox (dropdown) appearance for icon comboboxes
        self.option_add('*TCombobox*Listbox.background', '#2d2d2d')
        self.option_add('*TCombobox*Listbox.foreground', 'white')
        self.option_add('*TCombobox*Listbox.selectBackground', '#007ACC')
        self.option_add('*TCombobox*Listbox.selectForeground', 'white')

        # Red Cross Close
        if os.path.exists("icon2/close-window.png"):
             try:
                # Match Footer: 30x30, Red (#FF4444)
                # Use main_window's loader if available
                img = self.main_window.load_icon_colored("icon2/close-window.png", size=(30, 30), color="#FF4444")
                if img:
                    self.close_icon = img # Keep ref
                    btn_close = tk.Label(header, image=self.close_icon, bg=self.colors["bg_root"], cursor="hand2")
                else: 
                     raise Exception("Load failed")
             except Exception as e:
                print("Error loading Close icon: {}".format(e))
                btn_close = tk.Label(header, text="✕", fg="#FF4444", bg=self.colors["bg_root"], font=("Arial", 14, "bold"), cursor="hand2")
        else:
             btn_close = tk.Label(header, text="✕", fg="#FF4444", bg=self.colors["bg_root"], font=("Arial", 14, "bold"), cursor="hand2")
             
        btn_close.pack(side="right", padx=10)
        btn_close.bind("<Button-1>", lambda e: self.close_panel())

        # Attribution Info Button
        btn_info = tk.Label(header, text="ⓘ", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 12), cursor="hand2")
        btn_info.pack(side="right", padx=10)
        ToolTip(btn_info, "Icons made by IconKanan and Ardiansyah from www.flaticon.com", side="left")
        
        # Helper method to create section headers
        def create_section_header(parent, title, pady=(15, 5)):
            """Creates a section header with title and divider line."""
            section_frame = tk.Frame(parent, bg=self.colors["bg_root"])
            section_frame.pack(fill="x", padx=2, pady=pady)
            
            # Title label
            tk.Label(section_frame, text=title, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                    font=("Segoe UI", 11, "bold")).pack(side="left", anchor="w")
            
            # Divider line (partial width)
            divider = tk.Frame(section_frame, bg="#555555", height=1)
            divider.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # --- Scrollable Container ---
        self.scroll_frame = ScrollableFrame(self, bg=self.colors["bg_root"], auto_hide_scrollbar=False)
        self.scroll_frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        main_content = self.scroll_frame.scrollable_frame
        main_content.config(bg=self.colors["bg_root"])

        # === SECTION 1: Window Selection ===
        create_section_header(main_content, "Window Selection", pady=(2, 5))
        
        # --- Window Mode Selector ---
        window_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        window_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        # Track window mode (initialize from main window)
        self.window_mode_var = tk.StringVar(value=self.main_window.window_mode)
        
        # Determine initial button states
        is_single = (self.main_window.window_mode == "single")
        
        # Single Window Button
        self.btn_single_window = tk.Button(
            window_frame, text="Email Only", 
            command=lambda: self.select_window_mode("single"),
            bg=self.colors["accent"] if is_single else self.colors["bg_card"],
            fg="black" if is_single else "white",
            font=("Segoe UI", 10, "bold") if is_single else ("Segoe UI", 10),
            bd=0, padx=20, pady=4,
            activebackground=self.colors["accent"],
            activeforeground="black"
        )
        self.btn_single_window.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        # Dual Window Button
        self.btn_dual_window = tk.Button(
            window_frame, text="Emails & Reminders", 
            command=lambda: self.select_window_mode("dual"),
            bg=self.colors["accent"] if not is_single else self.colors["bg_card"],
            fg="black" if not is_single else "white",
            font=("Segoe UI", 10, "bold") if not is_single else ("Segoe UI", 10),
            bd=0, padx=20, pady=4,
            activebackground=self.colors["bg_card"],
            activeforeground="white"
        )
        self.btn_dual_window.pack(side="left", fill="x", expand=True)

        # === SECTION 2: General Settings ===
        create_section_header(main_content, "General Settings")

        # --- Typography Setting ---
        typo_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        typo_frame.pack(fill="x", padx=(20, 30), pady=(10, 0))
        
        tk.Label(typo_frame, text="Font Family:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        self.font_fam_cb = ttk.Combobox(typo_frame, values=["Segoe UI", "Arial", "Verdana", "Tahoma", "Courier New", "Georgia"], width=15, state="readonly", font=("Segoe UI", 10))
        self.font_fam_cb.set(self.main_window.font_family)
        self.font_fam_cb.pack(side="left", padx=(5, 20))
        self.font_fam_cb.bind("<<ComboboxSelected>>", self.update_font_settings)
        
        tk.Label(typo_frame, text="Size:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        
        # Use StringVar to ensure value is always visible
        self.font_size_var = tk.StringVar(value=str(self.main_window.font_size))
        self.font_size_cb = ttk.Combobox(
            typo_frame, 
            textvariable=self.font_size_var,
            values=[str(i) for i in range(8, 13)], 
            width=12, 
            state="readonly", 
            font=("Segoe UI", 10),
            style='FontSize.TCombobox'  # Use dedicated style
        )
        self.font_size_cb.pack(side="left", padx=5)
        self.font_size_cb.bind("<<ComboboxSelected>>", self.update_font_settings)
        
        # Proper postcommand to fix dropdown width and font
        def configure_font_size_dropdown():
            try:
                # Get the popdown window and its listbox
                popdown = self.font_size_cb.tk.call('ttk::combobox::PopdownWindow', self.font_size_cb)
                listbox = '{}.f.l'.format(popdown)
                
                # Set dropdown width to match or exceed combobox width
                cb_width = self.font_size_cb.winfo_width()
                min_width = max(cb_width, 100)  # At least 100 pixels
                self.font_size_cb.tk.call(listbox, 'configure', '-width', 20)  # 20 characters wide
                
                # Override font for THIS dropdown only (normal size, not the big icon font)
                self.font_size_cb.tk.call(listbox, 'configure', '-font', ('Segoe UI', 10))
            except:
                pass  # Silently fail if dropdown isn't ready
        
        self.font_size_cb['postcommand'] = configure_font_size_dropdown
        
        # --- System Settings (Refresh Rate) ---
        self.refresh_options = {"15s": 15, "30s": 30, "1m": 60, "2m": 120, "5m": 300}
        sys_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        sys_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        tk.Label(sys_frame, text="Refresh Rate:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        self.refresh_cb = ttk.Combobox(sys_frame, values=list(self.refresh_options.keys()), width=10, state="readonly", font=("Segoe UI", 10))
        
        current_label = "30s"
        for label, val in self.refresh_options.items():
            if val == self.main_window.poll_interval:
                current_label = label
                break
        self.refresh_cb.set(current_label)
        self.refresh_cb.pack(side="left", padx=5)
        self.refresh_cb.bind("<<ComboboxSelected>>", self.update_refresh_rate)

        # === SECTION 3: Email Settings ===
        create_section_header(main_content, "Email Settings")

        # Account Selection Button
        def open_drawer():
             self.close_panel() # Close settings
             self.main_window.toggle_account_selection() # Open drawer

        btn_accounts = tk.Button(main_content, text="Select Emails...", command=open_drawer,
                                 bg=self.colors["bg_card"], fg=self.colors["fg_text"], bd=0, font=("Segoe UI", 10),
                                 highlightthickness=1, highlightbackground=self.colors["divider"], pady=8)
        btn_accounts.pack(fill="x", padx=(18, 30), pady=(5, 5))

        # --- Email List Settings ---
        list_settings_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        list_settings_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        self.show_read_var = tk.BooleanVar(value=self.main_window.show_read)
        
        # Add trace callback
        def on_show_read_change(*args):
            self.update_email_filters()
        
        self.show_read_var.trace("w", on_show_read_change)
        
        self.chk_show_read = tk.Checkbutton(
            list_settings_frame, text="Include read email", 
            variable=self.show_read_var,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 10)
        )
        self.chk_show_read.grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.show_has_attachment_var = tk.BooleanVar(value=self.main_window.show_has_attachment)
        
        # Add trace callback
        def on_show_attachment_change(*args):
            self.update_email_filters()
        
        self.show_has_attachment_var.trace("w", on_show_attachment_change)
        
        self.chk_has_attachment = tk.Checkbutton(
            list_settings_frame, text="Show if has Attachment", 
            variable=self.show_has_attachment_var,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 10)
        )
        self.chk_has_attachment.grid(row=1, column=0, sticky="w", pady=(0, 5))
        
        # --- Email Window Content ---
        tk.Label(list_settings_frame, text="Email Window Content", 
                 bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                 font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(10, 2))
        
        self.email_content_frame = tk.Frame(list_settings_frame, bg=self.colors["bg_root"])
        self.email_content_frame.grid(row=3, column=0, sticky="w", padx=(20, 0))
        
        # Checkboxes
        self.email_show_sender_var = tk.BooleanVar(value=self.main_window.email_show_sender)
        tk.Checkbutton(self.email_content_frame, text="Who From", variable=self.email_show_sender_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
                       
        self.email_show_subject_var = tk.BooleanVar(value=self.main_window.email_show_subject)
        tk.Checkbutton(self.email_content_frame, text="Subject Line", variable=self.email_show_subject_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w")

        self.email_show_body_var = tk.BooleanVar(value=self.main_window.email_show_body)
        tk.Checkbutton(self.email_content_frame, text="Content Body", variable=self.email_show_body_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w")
        
        # Number of Lines Selector
        lines_frame = tk.Frame(self.email_content_frame, bg=self.colors["bg_root"])
        lines_frame.grid(row=3, column=0, sticky="w", pady=(5,0))
        tk.Label(lines_frame, text="Lines:", bg=self.colors["bg_root"], fg=self.colors["fg_secondary"], font=("Segoe UI", 10)).pack(side="left")
        
        self.email_body_lines_var = tk.StringVar(value=str(self.main_window.email_body_lines))
        self.cb_lines = ttk.Combobox(lines_frame, textvariable=self.email_body_lines_var, values=["1", "2", "3", "4"], width=3, state="readonly", font=("Segoe UI", 8))
        self.cb_lines.pack(side="left", padx=5)
        self.cb_lines.bind("<<ComboboxSelected>>", self.update_email_filters)
        
        # Configure dropdown font size
        def configure_lines_dropdown():
             try:
                 popdown = self.cb_lines.tk.call('ttk::combobox::PopdownWindow', self.cb_lines)
                 listbox = '{}.f.l'.format(popdown)
                 self.cb_lines.tk.call(listbox, 'configure', '-font', ('Segoe UI', 10))
             except:
                 pass
        self.cb_lines['postcommand'] = configure_lines_dropdown

        # "Show Content on Hover" - inside the content frame
        self.show_hover_content_var = tk.BooleanVar(value=self.main_window.show_hover_content)
        tk.Checkbutton(self.email_content_frame, text="Show Content on Hover", 
                       variable=self.show_hover_content_var, 
                       command=self.update_email_filters, 
                       bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], 
                       activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], 
                       font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w")

        # --- Interaction Settings (Merged into Email Settings) ---
        interaction_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        interaction_frame.pack(fill="x", padx=(18, 30), pady=(5, 10))
        
        self.buttons_on_hover_var = tk.BooleanVar(value=self.main_window.buttons_on_hover)
        tk.Checkbutton(interaction_frame, text="Show Buttons on Hover", variable=self.buttons_on_hover_var, 
                       command=self.update_interaction_settings, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], font=("Segoe UI", 9)).pack(side="left")
                       
        self.email_double_click_var = tk.BooleanVar(value=self.main_window.email_double_click)
        tk.Checkbutton(interaction_frame, text="Double Click to Open", variable=self.email_double_click_var, 
                       command=self.update_interaction_settings, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground=self.colors["fg_text"], font=("Segoe UI", 9)).pack(side="left", padx=10)

        create_section_header(main_content, "Quick Create")
        
        qc_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        qc_frame.pack(fill="x", padx=(18, 30), pady=(10, 10))
        
        self.qc_options = ["New Email", "New Meeting", "New Appointment", "New Task"]
        self.qc_vars = {}
        
        # Load current
        current_qc = getattr(self.main_window, "quick_create_actions", ["New Email"])
        
        def update_qc_settings():
            selected = [opt for opt, var in self.qc_vars.items() if var.get()]
            self.main_window.quick_create_actions = selected
            self.main_window.save_config()
            self.main_window.update_quick_create_icon()

        for idx, opt in enumerate(self.qc_options):
            var = tk.BooleanVar(value=(opt in current_qc))
            self.qc_vars[opt] = var
            chk = tk.Checkbutton(
                qc_frame, text=opt, variable=var,
                command=update_qc_settings,
                bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground=self.colors["fg_text"],
                font=("Segoe UI", 9)
            )
            chk.grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 15), pady=2)


        # === Button Configuration Table (Restored Original) ===
        # --- Button Configuration Table ---
        create_section_header(main_content, "Hover Buttons")
        
        container = tk.Frame(main_content, bg=self.colors["bg_root"], pady=12)
        container.pack(fill="x", expand=False, padx=(2, 20))  # 2px left padding
        
        # Table Headers
        headers = ["Icon", "Action", "Folders (for Move)"]
        
        for col, text in enumerate(headers):
            tk.Label(
                container, text=text, 
                bg=self.colors["bg_root"], fg=self.colors["fg_dim"], 
                font=("Segoe UI", 9)
            ).grid(row=0, column=col, sticky="w", padx=8, pady=(0, 8))
            
        # Rows
        self.rows_data = [] 
        self.action_options = ["None", "Mark Read", "Delete", "Read & Delete", "Flag", "Open Email", "Reply", "Move To..."]
        # Monochrome / Clean Unicode Icons AND Custom PNGs
        unicode_icons = [u"", u"🗑️", u"✉️", u"⚑", u"↩️", u"📂", u"↗", u"✓", u"✕", u"⚠"]
        
        # Scan for PNGs
        png_icons = []
        if os.path.exists("icons"):
            for file in glob.glob("icons/*.png"):
                png_icons.append(os.path.basename(file))
        
        self.icons = unicode_icons + png_icons
        
        # Auto-Icon Logic Map
        self.ACTION_TO_ICON = {
            "Reply": "Reply.png",
            "Delete": "Delete.png",
            "Mark Read": "Mark as Read.png",
            "Read & Delete": "Read & Delete.png",
            "Open Email": "open.png",
            "Flag": "Flag.png",
            "Move To...": "Move to Folder.png",
            "None": ""
        }
        
        current_config = self.main_window.btn_config
        row_config = current_config + [{}] * (4 - len(current_config))
        
        for i in range(4):
            c_data = row_config[i]
            
            # 1. Icon Display (Dynamic Label)
            lbl_icon = tk.Label(container, bg=self.colors["bg_root"], width=5) # Width roughly matches 30px
            lbl_icon.grid(row=i+1, column=0, padx=8, pady=5)
            
            # Preserve the icon value for saving (start with current)
            current_icon_val = c_data.get("icon", self.icons[0])
            
            # 2. Action (Previously Action 1)
            cb_act1 = ttk.Combobox(container, values=self.action_options, width=15, state="readonly", font=("Segoe UI", 10))
            cb_act1.set(c_data.get("action1", "None")) 
            cb_act1.grid(row=i+1, column=1, padx=8, pady=5, ipady=1)
            
            # 3. Folder Picker UI (Entry + Button) - Shifted to Column 2
            f_frame = tk.Frame(container, bg=self.colors["bg_root"])
            f_frame.grid(row=i+1, column=2, padx=8, pady=5)
            
            e_folder = ttk.Entry(f_frame, width=15, font=("Segoe UI", 10))
            e_folder.insert(0, c_data.get("folder", ""))
            e_folder.pack(side="left", ipady=1)
            e_folder.bind("<FocusOut>", lambda e: self.update_button_config())

            # Picker Button
            btn_pick = tk.Label(f_frame, text="...", bg=self.colors["bg_card"], fg=self.colors["fg_text"], font=("Segoe UI", 8), width=3, cursor="hand2")
            btn_pick.pack(side="left", padx=(5,0), fill="y")
            
            # Bind picker
            def open_picker(event, entry=e_folder):
                # Get folders from the first enabled email account (faster than all stores)
                account_name = None
                if self.main_window.enabled_accounts:
                    for name, conf in self.main_window.enabled_accounts.items():
                        if conf.get("email"):
                            account_name = name
                            break
                
                folders = self.main_window.outlook_client.get_folder_list(account_name)
                if not folders:
                    folders = ["Inbox"]
                
                # Show inline picker: hide settings scroll, show picker frame
                self.scroll_frame.pack_forget()
                
                def on_folder_selected(path):
                    # path may be a list from FolderPickerFrame
                    if isinstance(path, list):
                        path = path[0] if path else ""
                    entry.delete(0, tk.END)
                    entry.insert(0, path)
                    self.update_button_config()
                
                def on_cancel():
                    # Remove picker and restore settings
                    if hasattr(self, '_inline_picker') and self._inline_picker:
                        self._inline_picker.destroy()
                        self._inline_picker = None
                    self.scroll_frame.pack(fill="both", expand=True, padx=2, pady=2)
                
                self._inline_picker = FolderPickerFrame(
                    self, folders, on_folder_selected, on_cancel, 
                    selected_paths=[entry.get()] if entry.get() else None,
                    colors=self.colors
                )
                self._inline_picker.pack(fill="both", expand=True, padx=2, pady=2)

            btn_pick.bind("<Button-1>", open_picker)
            
            # Helper to update icon display based on action
            def update_icon_display(action_widget, icon_label, row_idx):
                action = action_widget.get()
                new_icon = self.ACTION_TO_ICON.get(action, "")
                
                # Update visual
                if new_icon:
                     # Check if PNG or Unicode
                     if new_icon.lower().endswith(".png"):
                         path = os.path.join("icons", new_icon)
                         if os.path.exists(path):
                             # Load using main_window's loader
                             img = self.main_window.load_icon_colored(path, size=(24, 24), color="#FFFFFF")
                             if img:
                                 # Keep reference to avoid GC
                                 setattr(icon_label, "image", img) 
                                 # IMPORTANT: Reset width to 0 (auto) when showing image, otherwise '5' means 5 pixels!
                                 icon_label.config(image=img, text="", width=0)
                             else:
                                 icon_label.config(text="?", image="", width=5)
                         else:
                             icon_label.config(text="?", image="", width=5)
                     else:
                         # Unicode
                         icon_label.config(text=new_icon, image="", fg="white", font=("Segoe UI", 16), width=5)
                else:
                    icon_label.config(text="", image="", width=5)
                
                if len(self.rows_data) > row_idx:
                    self.rows_data[row_idx]["icon_val"] = new_icon

            # Helper for visibility
            def update_folder_visibility(action_widget, folder_frame):
                 if action_widget.get() == "Move To...":
                      folder_frame.grid()
                 else:
                      folder_frame.grid_remove()

            # Auto-update Handler
            def on_action_change(event, act_cb=cb_act1, icon_lbl=lbl_icon, idx=i, f_frm=f_frame):
                 update_icon_display(act_cb, icon_lbl, idx)
                 update_folder_visibility(act_cb, f_frm)
                 self.refresh_dropdown_options() # Enforce uniqueness
                 self.update_button_config()  # Apply changes immediately
            
            cb_act1.bind("<<ComboboxSelected>>", on_action_change)
            
            self.rows_data.append({
                "icon_val": current_icon_val, # Store value directly
                "act1": cb_act1,
                "folder": e_folder,
                "folder_frame": f_frame
            })
            
            # Trigger initial display update manually
            update_icon_display(cb_act1, lbl_icon, i)
            update_folder_visibility(cb_act1, f_frame)
            
        # Initial Refresh of Options
        self.refresh_dropdown_options()

        # === SECTION 4: Reminder Settings ===
        create_section_header(main_content, "Reminder Settings")
        
        reminder_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        reminder_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        # --- 1. Follow-up Flags ---
        self.reminder_show_flagged_var = tk.BooleanVar(value=self.main_window.reminder_show_flagged)
        chk_followup = tk.Checkbutton(
            reminder_frame, text="Follow-up Flags", 
            variable=self.reminder_show_flagged_var,
            command=self.toggle_followup_options,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9, "bold")
        )
        chk_followup.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button for showing/hiding options
        self.followup_options_visible = False  # Start CLOSED

        # Unified Container for Hover Logic
        self.followup_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.followup_container.grid(row=0, column=1, sticky="nw", rowspan=2, padx=(5, 0))
        
        # Toggle button (Arrow) inside container
        self.btn_toggle_followup = tk.Label(
            self.followup_container, text="▼",
            bg=self.colors["bg_root"], fg=self.colors["fg_dim"],
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_followup.pack(side="top", anchor="w", pady=(2, 0))
        self.btn_toggle_followup.bind("<Button-1>", lambda e: self.toggle_followup_visibility())
        
        # Initially hide button if Follow-up Flags is unchecked
        if not self.main_window.reminder_show_flagged:
            self.followup_container.grid_remove() # Hide entire container
        
        # Container for due date checkboxes (conditionally shown)
        self.followup_options_frame = tk.Frame(self.followup_container, bg=self.colors["bg_root"])
        
        # Due date checkboxes
        self.due_options = ["Today", "Tomorrow", "This Week", "Next Week", "Overdue", "No Date"]
        self.due_vars = {}
        
        for idx, option in enumerate(self.due_options):
            var = tk.BooleanVar(value=False)  # Default to unchecked
            self.due_vars[option] = var
            
            chk = tk.Checkbutton(
                self.followup_options_frame, text=option,
                variable=var,
                command=self.update_reminder_filters,
                bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground=self.colors["fg_text"],
                font=("Segoe UI", 9)
            )
            # Use grid inside the packed frame
            chk.grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 10), pady=1) # 2 columns for compactness

        # "All" checkbox
        self.due_all_var = tk.BooleanVar(value=False)
        chk_all = tk.Checkbutton(
            self.followup_options_frame, text="All",
            variable=self.due_all_var,
            command=self.toggle_all_due_options,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9, "bold")
        )
        chk_all.grid(row=3, column=0, sticky="w", pady=(5, 0))

        # IMPORTANT: adjust layout of main checkbox to align
        chk_followup.grid(row=0, column=0, sticky="nw", pady=(0, 5)) 

        # Initially hide if not enabled
        if not self.main_window.reminder_show_flagged:
             self.followup_options_frame.pack_forget()
        
        # --- 2. Categories ---
        self.reminder_show_categorized_var = tk.BooleanVar(value=self.main_window.reminder_show_categorized)
        chk_categorized = tk.Checkbutton(
            reminder_frame, text="Categories", 
            variable=self.reminder_show_categorized_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9, "bold")
        )
        chk_categorized.grid(row=3, column=0, sticky="w", pady=(0, 5), columnspan=3)
        
        # --- 3. Importance ---
        self.reminder_show_importance_var = tk.BooleanVar(value=self.main_window.reminder_show_importance)  # Initialize from config
        chk_importance = tk.Checkbutton(
            reminder_frame, text="Importance", 
            variable=self.reminder_show_importance_var,
            command=self.toggle_importance_options,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9, "bold")
        )
        chk_importance.grid(row=4, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button for showing/hiding options
        self.importance_options_visible = False
        
        # Unified Container for Hover Logic
        self.importance_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.importance_container.grid(row=4, column=1, sticky="nw", rowspan=2, padx=(5, 0))

        self.btn_toggle_importance = tk.Label(
            self.importance_container, text="▼",
            bg=self.colors["bg_root"], fg=self.colors["fg_dim"],
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_importance.grid(row=0, column=0, sticky="w", pady=(2, 0))
        self.btn_toggle_importance.bind("<Button-1>", lambda e: self.toggle_importance_visibility())

        
        # Container for importance checkboxes
        self.importance_options_frame = tk.Frame(self.importance_container, bg=self.colors["bg_root"])

        # Adjust master checkbox alignment
        chk_importance.grid(row=4, column=0, sticky="nw", pady=(0, 5))
        
        self.reminder_high_importance_var = tk.BooleanVar(value=self.main_window.reminder_high_importance)
        chk_high = tk.Checkbutton(
            self.importance_options_frame, text="High", 
            variable=self.reminder_high_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9)
        )
        # Use grid inside packed frame
        chk_high.grid(row=0, column=0, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_normal_importance_var = tk.BooleanVar(value=self.main_window.reminder_normal_importance)
        chk_normal = tk.Checkbutton(
            self.importance_options_frame, text="Normal", 
            variable=self.reminder_normal_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9)
        )
        chk_normal.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_low_importance_var = tk.BooleanVar(value=self.main_window.reminder_low_importance)
        chk_low = tk.Checkbutton(
            self.importance_options_frame, text="Low", 
            variable=self.reminder_low_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9)
        )
        chk_low.grid(row=0, column=2, sticky="w", pady=2)
        
        # importance_options_frame starts hidden (not packed yet)
        
        # --- 4. Meetings ---
        self.reminder_show_meetings_var = tk.BooleanVar(value=self.main_window.reminder_show_meetings)  # Initialize from config
        chk_meetings = tk.Checkbutton(
            reminder_frame, text="Meetings", 
            variable=self.reminder_show_meetings_var,
            command=self.toggle_meetings_options,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9, "bold")
        )
        chk_meetings.grid(row=6, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button (Arrow)
        self.meetings_options_visible = False
        
        # Unified Container
        self.meetings_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.meetings_container.grid(row=6, column=1, sticky="nw", rowspan=2, padx=(5, 0))

        self.btn_toggle_meetings = tk.Label(
            self.meetings_container, text="▼",
            bg=self.colors["bg_root"], fg=self.colors["fg_dim"],
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_meetings.grid(row=0, column=0, sticky="w", pady=(2, 0))
        self.btn_toggle_meetings.bind("<Button-1>", lambda e: self.toggle_meetings_visibility())

        
        # Container for meeting status options
        self.meetings_options_frame = tk.Frame(self.meetings_container, bg=self.colors["bg_root"])

        # NEW: Single Column Layout for Meetings
        self.meeting_statuses = ["Accepted", "Tentative", "Appointments", "Received/Unknown"]
        self.meeting_vars = {}
        
        # Map nice names to internal status
        # 3=Accepted, 2=Tentative, 0=None (Appointment)
        
        for idx, status in enumerate(self.meeting_statuses):
            var = tk.BooleanVar(value=True) # Default all on
            self.meeting_vars[status] = var
            chk = tk.Checkbutton(
                self.meetings_options_frame, text=status,
                variable=var,
                command=self.update_reminder_filters,
                bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground=self.colors["fg_text"],
                font=("Segoe UI", 9)
            )
            # Use vertical packing (grid with 1 col)
            chk.grid(row=idx, column=0, sticky="w", padx=(0, 10), pady=2)
            
        # meetings_options_frame starts hidden (not packed yet)

        # === TASKS (New Section) ===
        # --- 5. Tasks ---
        self.reminder_show_tasks_var = tk.BooleanVar(value=self.main_window.reminder_show_tasks)
        chk_tasks = tk.Checkbutton(
            reminder_frame, text="Tasks", 
            variable=self.reminder_show_tasks_var,
            command=self.toggle_tasks_options,
            bg=self.colors["bg_root"], fg=self.colors["fg_text"],
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground=self.colors["fg_text"],
            font=("Segoe UI", 9, "bold")
        )
        chk_tasks.grid(row=8, column=0, sticky="w", pady=(0, 5))

        # Toggle button
        self.tasks_options_visible = False
        self.tasks_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.tasks_container.grid(row=8, column=1, sticky="nw", rowspan=2, padx=(5, 0))

        self.btn_toggle_tasks = tk.Label(
            self.tasks_container, text="▼",
            bg=self.colors["bg_root"], fg=self.colors["fg_dim"],
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_tasks.grid(row=0, column=0, sticky="w", pady=(2, 0))
        self.btn_toggle_tasks.bind("<Button-1>", lambda e: self.toggle_tasks_visibility())


        self.tasks_options_frame = tk.Frame(self.tasks_container, bg=self.colors["bg_root"])
        # Vertical Layout for Task Types + Date Filters
        
        # Task Types
        self.task_types = ["Tasks", "To-Do"] # Only real tasks for now
        self.task_vars = {}
        for idx, t in enumerate(self.task_types):
             var = tk.BooleanVar(value=True)
             self.task_vars[t] = var
             tk.Checkbutton(
                 self.tasks_options_frame, text=t, variable=var,
                 command=self.update_reminder_filters,
                 bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                 selectcolor=self.colors["bg_card"],
                 font=("Segoe UI", 9)
             ).grid(row=idx, column=0, sticky="w", pady=1)

        # Separator
        tk.Frame(self.tasks_options_frame, bg="#444444", height=1).grid(row=2, column=0, sticky="ew", pady=5)

        # Date Filters for Tasks
        self.task_date_options = ["Today", "Tomorrow", "Next 7 Days", "No Date", "Overdue"]
        self.task_date_vars = {}
        
        for idx, option in enumerate(self.task_date_options):
            var = tk.BooleanVar(value=False)
            self.task_date_vars[option] = var
            tk.Checkbutton(
                 self.tasks_options_frame, text=option,
                 variable=var,
                 command=self.update_reminder_filters,
                 bg=self.colors["bg_root"], fg=self.colors["fg_text"],
                 selectcolor=self.colors["bg_card"],
                 font=("Segoe UI", 9)
            ).grid(row=3+idx, column=0, sticky="w", pady=1)

        # tasks_options_frame starts hidden (not gridded yet)

        # Update tick states for meetings/tasks
        self.update_meeting_ticks_from_config()
        self.update_task_ticks_from_config()

        # Update visibility states based on initial vars
        if self.reminder_show_importance_var.get():
             self.btn_toggle_importance.grid()
        if self.reminder_show_meetings_var.get():
             self.btn_toggle_meetings.grid()
        if self.reminder_show_tasks_var.get():
             self.btn_toggle_tasks.grid()

    def update_interaction_settings(self):
        self.main_window.buttons_on_hover = self.buttons_on_hover_var.get()
        self.main_window.email_double_click = self.email_double_click_var.get()
        self.main_window.save_config()
        self.main_window.refresh_emails()

    def toggle_email_content_options(self):
        if self.email_content_visible:
            self.email_content_frame.grid_remove()
            self.email_content_visible = False
        else:
            self.email_content_frame.grid()
            self.email_content_visible = True

    def update_email_filters(self, *args):
        # Update Main Window config
        self.main_window.show_read = self.show_read_var.get()
        self.main_window.show_has_attachment = self.show_has_attachment_var.get()
        self.main_window.email_show_sender = self.email_show_sender_var.get()
        self.main_window.email_show_subject = self.email_show_subject_var.get()
        self.main_window.email_show_body = self.email_show_body_var.get()
        self.main_window.show_hover_content = self.show_hover_content_var.get()
        
        try:
             lines = int(self.email_body_lines_var.get())
             self.main_window.email_body_lines = lines
        except: pass
        
        self.main_window.save_config()
        self.main_window.refresh_emails()
        
    def update_font_settings(self, event=None):
        fam = self.font_fam_cb.get()
        try:
            size = int(self.font_size_cb.get())
        except:
            size = 9
            
        self.main_window.font_family = fam
        self.main_window.font_size = size
        self.main_window.save_config()
        
        # Reload UI fonts? Requires restart or huge refresh.
        # For now, just save.
        messagebox.showinfo("Fonts Updated", "Font settings saved. Please restart the app for changes to apply fully.")

    def update_refresh_rate(self, event=None):
        label = self.refresh_cb.get()
        val = self.refresh_options.get(label, 30)
        self.main_window.poll_interval = val
        self.main_window.save_config()

    def select_window_mode(self, mode):
        self.main_window.window_mode = mode
        self.window_mode_var.set(mode)
        self.main_window.save_config()
        
        # Update button visuals
        is_single = (mode == "single")
        self.btn_single_window.config(
            bg=self.colors["accent"] if is_single else self.colors["bg_card"],
            fg="black" if is_single else "white",
            font=("Segoe UI", 10, "bold") if is_single else ("Segoe UI", 10)
        )
        self.btn_dual_window.config(
            bg=self.colors["accent"] if not is_single else self.colors["bg_card"],
            fg="black" if not is_single else "white",
            font=("Segoe UI", 10, "bold") if not is_single else ("Segoe UI", 10)
        )
        
        # Trigger Resize/Reflow
        self.main_window.apply_window_mode()

    def refresh_dropdown_options(self):
        """Refreshes dropdown options to discourage duplicates."""
        # Get all currently selected actions
        selected_actions = []
        for row in self.rows_data:
             val = row["act1"].get()
             if val != "None":
                 selected_actions.append(val)
        
        # For each row, rebuild values list
        # We allow a value if it's currently selected in THIS row, OR not selected anywhere else.
        # BUT user wants to swap, so maybe just allow all but warn?
        # Actually, let's just stick to allowing all for flexibility, 
        # but maybe obscure ones already used? No, duplicate actions might be desired (e.g. Delete on left and right)
        pass

    def update_button_config(self, event=None):
        new_config = []
        for row in self.rows_data:
             entry = {
                 "icon": row["icon_val"], # Persist the logic-determined icon
                 "action1": row["act1"].get(),
                 "folder": row["folder"].get()
             }
             new_config.append(entry)
        
        self.main_window.btn_config = new_config
        self.main_window.save_config()
        self.main_window.refresh_emails() # Redraw buttons

    def close_panel(self):
        self.main_window.toggle_settings_panel()

    # --- Methods for Reminder Settings Logic ---
    def toggle_followup_options(self):
        if self.reminder_show_flagged_var.get():
             self.main_window.reminder_show_flagged = True
             self.btn_toggle_followup.grid() # Show arrow
             self.toggle_followup_visibility(force_open=True)
        else:
             self.main_window.reminder_show_flagged = False
             self.btn_toggle_followup.grid_remove()
             self.followup_options_frame.pack_forget() # Hide options
        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def toggle_followup_visibility(self, force_open=False):
        if self.followup_options_visible and not force_open:
            self.followup_options_frame.pack_forget()
            self.btn_toggle_followup.config(text="▼")
            self.followup_options_visible = False
        else:
            # Accordion: close other sections
            self._close_other_accordions('followup')
            self.followup_options_frame.pack(side="top", anchor="w", padx=(10, 0))
            self.btn_toggle_followup.config(text="▲")
            self.followup_options_visible = True
            self.after(50, lambda: self._scroll_into_view(self.followup_options_frame))

    def toggle_all_due_options(self):
        state = self.due_all_var.get()
        # Set all individual vars
        for var in self.due_vars.values():
             var.set(state)
        self.update_reminder_filters()

    def toggle_importance_options(self):
        if self.reminder_show_importance_var.get():
             self.main_window.reminder_show_importance = True
             self.importance_container.grid()
             self.toggle_importance_visibility(force_open=True)
        else:
             self.main_window.reminder_show_importance = False
             self.importance_options_frame.grid_remove()
             self.importance_container.grid_remove()
        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def toggle_importance_visibility(self, force_open=False):
        if self.importance_options_visible and not force_open:
             self.importance_options_frame.grid_remove()
             self.btn_toggle_importance.config(text="▼")
             self.importance_options_visible = False
        else:
             # Accordion: close other sections
             self._close_other_accordions('importance')
             self.importance_options_frame.grid(row=1, column=0, sticky="w", padx=(10, 0))
             self.btn_toggle_importance.config(text="▲")
             self.importance_options_visible = True
             self.after(50, lambda: self._scroll_into_view(self.importance_options_frame))

    def toggle_meetings_options(self):
        if self.reminder_show_meetings_var.get():
             self.main_window.reminder_show_meetings = True
             self.meetings_container.grid()
             self.toggle_meetings_visibility(force_open=True)
        else:
             self.main_window.reminder_show_meetings = False
             self.meetings_options_frame.grid_remove()
             self.meetings_container.grid_remove()
        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def toggle_meetings_visibility(self, force_open=False):
        if self.meetings_options_visible and not force_open:
             self.meetings_options_frame.grid_remove()
             self.btn_toggle_meetings.config(text="▼")
             self.meetings_options_visible = False
        else:
             # Accordion: close other sections
             self._close_other_accordions('meetings')
             self.meetings_options_frame.grid(row=1, column=0, sticky="w", padx=(10, 0))
             self.btn_toggle_meetings.config(text="▲")
             self.meetings_options_visible = True
             self.after(50, lambda: self._scroll_into_view(self.meetings_options_frame))

    def toggle_tasks_options(self):
        if self.reminder_show_tasks_var.get():
             self.main_window.reminder_show_tasks = True
             self.tasks_container.grid()
             self.toggle_tasks_visibility(force_open=True)
        else:
             self.main_window.reminder_show_tasks = False
             self.tasks_options_frame.grid_remove()
             self.tasks_container.grid_remove()
        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def toggle_tasks_visibility(self, force_open=False):
        if self.tasks_options_visible and not force_open:
             self.tasks_options_frame.grid_remove()
             self.btn_toggle_tasks.config(text="▼")
             self.tasks_options_visible = False
        else:
             # Accordion: close other sections
             self._close_other_accordions('tasks')
             self.tasks_options_frame.grid(row=1, column=0, sticky="w", padx=(10, 0))
             self.btn_toggle_tasks.config(text="▲")
             self.tasks_options_visible = True
             self.after(50, lambda: self._scroll_into_view(self.tasks_options_frame))

    def _close_other_accordions(self, except_section):
        """Close all accordion sections except the one being opened."""
        if except_section != 'followup' and self.followup_options_visible:
            self.followup_options_frame.pack_forget()
            self.btn_toggle_followup.config(text="▼")
            self.followup_options_visible = False
        if except_section != 'importance' and self.importance_options_visible:
            self.importance_options_frame.grid_remove()
            self.btn_toggle_importance.config(text="▼")
            self.importance_options_visible = False
        if except_section != 'meetings' and self.meetings_options_visible:
            self.meetings_options_frame.grid_remove()
            self.btn_toggle_meetings.config(text="▼")
            self.meetings_options_visible = False
        if except_section != 'tasks' and self.tasks_options_visible:
            self.tasks_options_frame.grid_remove()
            self.btn_toggle_tasks.config(text="▼")
            self.tasks_options_visible = False

    def _scroll_into_view(self, widget):
        """Scroll the settings panel so widget is visible."""
        try:
            self.update_idletasks()  # Force geometry update
            canvas = self.scroll_frame.canvas
            # Get widget position relative to the scrollable frame
            widget_y = widget.winfo_rooty() - self.scroll_frame.scrollable_frame.winfo_rooty()
            widget_h = widget.winfo_height()
            widget_bottom = widget_y + widget_h
            # Get visible area
            canvas_h = canvas.winfo_height()
            scroll_top = canvas.canvasy(0)
            scroll_bottom = scroll_top + canvas_h
            # If widget bottom is below visible area, scroll down
            if widget_bottom > scroll_bottom:
                # Scroll so the widget bottom is at the canvas bottom
                total_h = self.scroll_frame.scrollable_frame.winfo_height()
                if total_h > 0:
                    target = (widget_bottom - canvas_h) / total_h
                    canvas.yview_moveto(min(target, 1.0))
        except Exception:
            pass

    def update_reminder_filters(self):
        # Gather all filter states and save to main_window config
        
        # 1. Follow-up Dates
        active_due = [opt for opt, var in self.due_vars.items() if var.get()]
        self.main_window.reminder_followup_dates = active_due
        
        # Check "All" state
        if len(active_due) == len(self.due_options):
             self.due_all_var.set(True)
        else:
             self.due_all_var.set(False)

        # 2. Importance
        self.main_window.reminder_high_importance = self.reminder_high_importance_var.get()
        self.main_window.reminder_normal_importance = self.reminder_normal_importance_var.get()
        self.main_window.reminder_low_importance = self.reminder_low_importance_var.get()
        
        # 3. Meetings
        active_meetings = [opt for opt, var in self.meeting_vars.items() if var.get()]
        self.main_window.reminder_meeting_states = active_meetings
        
        # 4. Tasks
        active_task_dates = [opt for opt, var in self.task_date_vars.items() if var.get()]
        self.main_window.reminder_task_dates = active_task_dates

        self.main_window.reminder_show_categorized = self.reminder_show_categorized_var.get()

        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def update_meeting_ticks_from_config(self):
        current = self.main_window.reminder_meeting_states
        for status, var in self.meeting_vars.items():
             if status in current:
                 var.set(True)
             else:
                 var.set(False)

    def update_task_ticks_from_config(self):
        # Update Task Dates
        current_dates = self.main_window.reminder_task_dates
        for option, var in self.task_date_vars.items():
            if option in current_dates:
                 var.set(True)
            else:
                 var.set(False)
        
        # Update Due Checks (Follow-ups) if redundant? No, strictly for tasks now.
        pass
