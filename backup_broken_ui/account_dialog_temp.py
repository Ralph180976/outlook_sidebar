
class AccountSelectionDialog(tk.Toplevel):
    def __init__(self, parent, accounts, current_enabled, callback):
        super().__init__(parent)
        self.callback = callback
        self.accounts = accounts # List of account names
        
        # Ensure deep copy or at least working copy of settings
        # config structure: "Account Name": {"email": True, "calendar": False}
        self.working_settings = {}
        for acc in accounts:
            if acc in current_enabled:
                self.working_settings[acc] = current_enabled[acc].copy()
            else:
                # Default to True/True for new accounts
                self.working_settings[acc] = {"email": True, "calendar": True}

        # Win11 Colors (Dark Mode)
        self.colors = {
            "bg": "#202020",
            "fg": "#FFFFFF",
            "accent": "#60CDFF", 
            "secondary": "#444444",
            "border": "#2b2b2b"
        }
        
        self.title("Enabled Accounts")
        self.overrideredirect(True)
        self.wm_attributes("-topmost", True)
        self.config(bg=self.colors["bg"])
        self.configure(highlightbackground=self.colors["accent"], highlightthickness=1)
        
        # Geometry
        w, h = 400, 350
        x = parent.winfo_x() + 50
        y = parent.winfo_y() + 50
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        # --- Header ---
        header = tk.Frame(self, bg=self.colors["bg"], height=40)
        header.pack(fill="x", side="top")
        header.bind("<Button-1>", self.start_move)
        header.bind("<B1-Motion>", self.on_move)
        
        lbl = tk.Label(header, text="Select Accounts", bg=self.colors["bg"], fg=self.colors["fg"], 
                       font=("Segoe UI", 11, "bold"))
        lbl.pack(side="left", padx=15, pady=10)
        
        btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2", font=("Segoe UI", 10))
        btn_close.pack(side="right", padx=15)
        btn_close.bind("<Button-1>", lambda e: self.destroy())

        # --- Content Area (Scrollable) ---
        container = tk.Frame(self, bg=self.colors["bg"])
        container.pack(fill="both", expand=True, padx=2, pady=2)
        
        canvas = tk.Canvas(container, bg=self.colors["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=self.colors["bg"])
        
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- List Generation ---
        # Headers
        h_frame = tk.Frame(scroll_frame, bg=self.colors["bg"])
        h_frame.pack(fill="x", pady=(5, 10), padx=10)
        
        tk.Label(h_frame, text="Account", bg=self.colors["bg"], fg="#AAAAAA", width=25, anchor="w").pack(side="left")
        tk.Label(h_frame, text="Email", bg=self.colors["bg"], fg="#AAAAAA", width=8).pack(side="left")
        tk.Label(h_frame, text="Cal/Task", bg=self.colors["bg"], fg="#AAAAAA", width=8).pack(side="left")
        
        tk.Frame(scroll_frame, bg="#333333", height=1).pack(fill="x", padx=10, pady=(0, 5))

        self.vars = {} # {acc_name: {"email": IntVar, "calendar": IntVar}}
        
        for acc in self.accounts:
            row = tk.Frame(scroll_frame, bg=self.colors["bg"])
            row.pack(fill="x", padx=10, pady=2)
            
            # Truncate long names
            disp_name = acc if len(acc) < 30 else acc[:27] + "..."
            tk.Label(row, text=disp_name, bg=self.colors["bg"], fg="white", 
                     width=25, anchor="w", font=("Segoe UI", 9)).pack(side="left")
            
            self.vars[acc] = {}
            
            # Email Checkbox
            e_var = tk.IntVar(value=1 if self.working_settings[acc].get("email") else 0)
            self.vars[acc]["email"] = e_var
            # Custom Checkbox appearance is hard, using standard for now
            cb_e = tk.Checkbutton(row, variable=e_var, bg=self.colors["bg"], 
                                  activebackground=self.colors["bg"], selectcolor="#333333")
            cb_e.pack(side="left", padx=(10, 15))
            
            # Calendar Checkbox
            c_var = tk.IntVar(value=1 if self.working_settings[acc].get("calendar") else 0)
            self.vars[acc]["calendar"] = c_var
            cb_c = tk.Checkbutton(row, variable=c_var, bg=self.colors["bg"], 
                                  activebackground=self.colors["bg"], selectcolor="#333333")
            cb_c.pack(side="left", padx=10)
            
        # --- Footer Actions ---
        footer = tk.Frame(self, bg=self.colors["bg"], height=50)
        footer.pack(fill="x", side="bottom", pady=10)
        
        btn_save = tk.Button(footer, text="Save Changes", command=self.save_selection,
            bg=self.colors["accent"], fg="black", bd=0, font=("Segoe UI", 9, "bold"), padx=20, pady=5)
        btn_save.pack(side="right", padx=15)
        
        btn_cancel = tk.Button(footer, text="Cancel", command=self.destroy,
            bg="#333333", fg="white", bd=0, font=("Segoe UI", 9), padx=15, pady=5)
        btn_cancel.pack(side="right", padx=5)

    def save_selection(self):
        final_settings = {}
        for acc in self.accounts:
            final_settings[acc] = {
                "email": bool(self.vars[acc]["email"].get()),
                "calendar": bool(self.vars[acc]["calendar"].get())
            }
        self.callback(final_settings)
        self.destroy()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def on_move(self, event):
        deltax = event.x - self._x
        deltay = event.y - self._y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry(f"+{x}+{y}")
