# -*- coding: utf-8 -*-
from sidebar.core.compat import tk, ttk, messagebox
from PIL import Image, ImageTk
import os

from sidebar.core.config import RESAMPLE_MODE

class FolderPickerFrame(tk.Frame):
    def __init__(self, parent, folders, callback, on_cancel, selected_paths=None, colors=None):
        tk.Frame.__init__(self, parent)
        self.callback = callback
        self.on_cancel = on_cancel
        self.folders = folders
        self.selected_paths = set(selected_paths) if selected_paths else set()
        
        # Use provided colors or fallback to Win11 Dark
        if colors:
             self.colors = colors
             # Ensure keys exist or map them?
             # Mapping for compatibility if keys differ
             if "bg_root" in colors and "bg" not in colors:
                 self.colors["bg"] = colors["bg_root"]
                 self.colors["fg"] = colors["fg_text"]
                 self.colors["select_bg"] = colors["bg_card"] # or accent?
                 self.colors["dim"] = colors["fg_dim"]
        else:
            self.colors = {
                "bg": "#202020",
                "fg": "#FFFFFF",
                "accent": "#60CDFF", 
                "select_bg": "#444444",
                "dim": "#AAAAAA"
            }
        
        self.config(bg=self.colors["bg"])
        
        # Title Bar / Header
        header = tk.Frame(self, bg=self.colors["bg"])
        header.pack(fill="x", side="top", pady=(10, 5))

        lbl = tk.Label(header, text="Select Folder", bg=self.colors["bg"], fg=self.colors["fg"], font=("Segoe UI", 11, "bold"))
        lbl.pack(side="left", padx=15)
        
        # Close Button (Back/Cancel)
        if os.path.exists("icon2/close-window.png"):
             try:
                pil_img = Image.open("icon2/close-window.png").convert("RGBA")
                pil_img = pil_img.resize((20, 20), RESAMPLE_MODE)
                self.close_icon = ImageTk.PhotoImage(pil_img)
                btn_close = tk.Label(header, image=self.close_icon, bg=self.colors["bg"], cursor="hand2")
             except:
                btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg=self.colors["fg"], cursor="hand2")
        else:
             btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg=self.colors["fg"], cursor="hand2")

        btn_close.pack(side="right", padx=15)
        btn_close.bind("<Button-1>", lambda e: self.on_cancel())

        # Help Text (Under Header)
        tk.Label(self, text="Select folders to sync. Hold Shift/Ctrl for multiple.", 
                 bg=self.colors["bg"], fg=self.colors["dim"], font=("Segoe UI", 8)).pack(side="top", anchor="w", padx=15, pady=(0, 10))

        
        # Select Button (Packed at bottom FIRST so it stays visible)
        btn_sel = tk.Button(self, text="Save Selection", command=self.select_folder,
            bg=self.colors["accent"], fg="black", bd=0, font=("Segoe UI", 9, "bold"), pady=5)
        btn_sel.pack(side="bottom", fill="x", padx=10, pady=10)

        # TreeView (Packed LAST to fill remaining space)
        tree_frame = tk.Frame(self, bg=self.colors["bg"])
        tree_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
            background=self.colors.get("bg_card", "#2D2D30"), 
            foreground=self.colors.get("fg", "white"), 
            fieldbackground=self.colors.get("bg_card", "#2D2D30"),
            borderwidth=0
        )
        style.map("Treeview", background=[("selected", self.colors["accent"])])

        self.tree = ttk.Treeview(tree_frame, show="tree", selectmode="extended")
        self.tree.pack(side="left", fill="both", expand=True)
        
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        
        self.populate_tree()

    def populate_tree(self):
        nodes = {}
        for path in self.folders:
            parts = path.split("/")
            parent = ""
            current = ""
            for i, part in enumerate(parts):
                current = "{}/{}".format(parent, part) if parent else part
                if current not in nodes:
                    pid = parent if parent else ""
                    try:
                        nodes[current] = self.tree.insert(pid, "end", iid=current, text=part, open=False)
                    except: pass
                parent = current

        # Apply Selection
        if self.selected_paths:
            to_select = []
            for path in self.selected_paths:
                if self.tree.exists(path):
                    to_select.append(path)
                    # Open parents
                    # Simple walk up since we use path strings as iids
                    parts = path.split("/")
                    curr = ""
                    for p in parts[:-1]:
                        curr = "{}/{}".format(curr, p) if curr else p
                        if self.tree.exists(curr):
                            self.tree.item(curr, open=True)
            
            if to_select:
                try:
                    self.tree.selection_set(to_select)
                    self.tree.see(to_select[0])
                except: pass

    def select_folder(self):
        sel_items = self.tree.selection()
        if sel_items:
            paths = list(sel_items)
            self.callback(paths)
        # Always close/return on button press (even if nothing selected? logic implies save)
        self.on_cancel()


class FolderPickerWindow(tk.Toplevel):
    def __init__(self, parent, folders, callback, selected_paths=None, colors=None):
        tk.Toplevel.__init__(self, parent)
        self.callback = callback
        self.title("Select Folders")
        self.overrideredirect(True) 
        
        # Colors
        self.colors = colors if colors else {
            "bg_root": "#202020", "accent": "#60CDFF", "bg_card": "#2D2D30", "fg_text": "#FFFFFF", "fg_dim": "#AAAAAA"
        }
        
        # Config
        bg = self.colors.get("bg_root", "#202020")
        accent = self.colors.get("accent", "#60CDFF")
        
        self.config(bg=bg, highlightbackground=accent, highlightthickness=1)
        
        w, h = 350, 450
        # Center relative to parent if possible
        try:
            x = parent.winfo_x() + 60
            y = parent.winfo_y() + 60
        except:
            x, y = 100, 100
        self.geometry("{}x{}+{}+{}".format(w, h, x, y))
        
        def on_cancel():
            self.destroy()
            
        def on_done(val):
            callback(val)
            self.destroy()

        self.picker = FolderPickerFrame(self, folders, on_done, on_cancel, selected_paths, colors=self.colors)
        self.picker.pack(fill="both", expand=True, padx=2, pady=2)
        
        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.on_move)

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def on_move(self, event):
        deltax = event.x - self._x
        deltay = event.y - self._y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry("+{}+{}".format(x, y))


class AccountSelectionUI(tk.Frame):
    def __init__(self, parent, accounts, current_enabled, folder_selector, bg_color="#202020"):
        tk.Frame.__init__(self, parent, bg=bg_color)
        self.accounts = accounts
        self.current_enabled = current_enabled or {}
        self.folder_selector = folder_selector # Function(account, on_selected_callback)
        self.colors = {
            "bg": bg_color, "fg": "#FFFFFF", "accent": "#60CDFF", 
            "secondary": "#444444", "border": "#2b2b2b"
        }
        
        self.working_settings = {}
        for acc in accounts:
            if acc in self.current_enabled:
                self.working_settings[acc] = self.current_enabled[acc].copy()
            else:
                self.working_settings[acc] = {"email": True, "calendar": True}
                
        self.vars = {}
        self.setup_ui()
        
    def setup_ui(self):
        # --- Header with Close Button ---
        header = tk.Frame(self, bg=self.colors["bg"])
        header.pack(fill="x", pady=(10, 5))
        
        lbl_title = tk.Label(header, text="Select Accounts", bg=self.colors["bg"], fg="white", font=("Segoe UI", 11, "bold"))
        lbl_title.pack(side="left", padx=15)
        
        # Close Button
        if os.path.exists("icon2/close-window.png"):
             try:
                pil_img = Image.open("icon2/close-window.png").convert("RGBA")
                pil_img = pil_img.resize((20, 20), RESAMPLE_MODE)
                self.close_icon = ImageTk.PhotoImage(pil_img)
                btn_close = tk.Label(header, image=self.close_icon, bg=self.colors["bg"], cursor="hand2")
             except:
                btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")
        else:
             btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")

        btn_close.pack(side="right", padx=15)
        # Assuming parent is the overlay frame calling .place()
        btn_close.bind("<Button-1>", lambda e: self.master.place_forget())

        # Help text removed - the FolderPickerFrame shows its own hint when the folder tree appears
        
        # Scrollable Area
        canvas = tk.Canvas(self, bg=self.colors["bg"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=self.colors["bg"])
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        def on_canvas_configure(event):
            canvas.itemconfig(win_id, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)
        
        canvas.configure(yscrollcommand=self.scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Headers
        h_frame = tk.Frame(scroll_frame, bg=self.colors["bg"])
        h_frame.pack(fill="x", pady=(5, 10), padx=5)
        tk.Label(h_frame, text="Account", bg=self.colors["bg"], fg="#AAAAAA", width=20, anchor="w", font=("Segoe UI", 9)).pack(side="left")
        tk.Label(h_frame, text="Mail", bg=self.colors["bg"], fg="#AAAAAA", width=4, font=("Segoe UI", 9)).pack(side="left")
        tk.Label(h_frame, text="Fldr", bg=self.colors["bg"], fg="#AAAAAA", width=4, font=("Segoe UI", 9)).pack(side="left")
        tk.Label(h_frame, text="Cal", bg=self.colors["bg"], fg="#AAAAAA", width=4, font=("Segoe UI", 9)).pack(side="left")

        tk.Frame(scroll_frame, bg="#333333", height=1).pack(fill="x", padx=5, pady=(0, 5))

        # Preload folder icon
        self.folder_icon_img = None
        if os.path.exists("icon2/folder.png"):
             try:
                pil_img = Image.open("icon2/folder.png").convert("RGBA")
                
                # Create a solid white image of the same size
                white_img = Image.new("RGBA", pil_img.size, (255, 255, 255, 255))
                # Use the alpha channel from the original image as a mask/channel
                r, g, b, a = pil_img.split()
                white_img.putalpha(a)
                pil_img = white_img
                
                pil_img = pil_img.resize((20, 20), RESAMPLE_MODE)
                self.folder_icon_img = ImageTk.PhotoImage(pil_img)
             except Exception as e:
                print("Error loading folder icon: {}".format(e))

        for acc in self.accounts:
            row = tk.Frame(scroll_frame, bg=self.colors["bg"])
            row.pack(fill="x", padx=5, pady=2)
            
            disp = acc if len(acc) < 25 else acc[:22] + "..."
            tk.Label(row, text=disp, bg=self.colors["bg"], fg="white", 
                     width=20, anchor="w", font=("Segoe UI", 9)).pack(side="left")
            
            self.vars[acc] = {}
            vals = self.working_settings[acc]
            
            # Email Check
            e_var = tk.IntVar(value=1 if vals.get("email") else 0)
            self.vars[acc]["email"] = e_var
            # "Pale Gray" checkbox #CCCCCC. Font size hack for size.
            tk.Checkbutton(row, variable=e_var, bg=self.colors["bg"], activebackground=self.colors["bg"], 
                           selectcolor="#CCCCCC", activeforeground="white", borderwidth=0, highlightthickness=0,
                           font=("Segoe UI", 14)).pack(side="left", padx=(5, 5))
            
            # Folder Button
            self.vars[acc]["email_folders"] = vals.get("email_folders", [])
            
            if self.folder_icon_img:
                btn_f = tk.Label(row, image=self.folder_icon_img, bg=self.colors["bg"], cursor="hand2")
            else:
                btn_f = tk.Label(row, text="📁", bg=self.colors["bg"], fg=self.colors["accent"], cursor="hand2", font=("Segoe UI", 10))
                
            btn_f.pack(side="left", padx=10)
            btn_f.bind("<Button-1>", lambda e, a=acc: self.on_folder_click(a))

            # Calendar Check
            c_var = tk.IntVar(value=1 if vals.get("calendar") else 0)
            self.vars[acc]["calendar"] = c_var
            tk.Checkbutton(row, variable=c_var, bg=self.colors["bg"], activebackground=self.colors["bg"], 
                           selectcolor="#CCCCCC", activeforeground="white", borderwidth=0, highlightthickness=0,
                           font=("Segoe UI", 14)).pack(side="left", padx=5)

    def on_folder_click(self, account):
        def on_selected(paths):
            self.vars[account]["email_folders"] = paths
            
        current_paths = self.vars[account]["email_folders"]
        self.folder_selector(account, on_selected, current_paths)

    def get_settings(self):
        final = {}
        for acc in self.accounts:
            final[acc] = {
                "email": bool(self.vars[acc]["email"].get()),
                "calendar": bool(self.vars[acc]["calendar"].get()),
                "email_folders": self.vars[acc]["email_folders"]
            }
        return final

class AccountSelectionDialog(tk.Toplevel):
    def __init__(self, parent, accounts, current_enabled, callback, colors=None):
        tk.Toplevel.__init__(self, parent)
        self.callback = callback
        self.colors = colors if colors else {
            "bg": "#202020", "fg": "#FFFFFF", "accent": "#60CDFF"
        }
        
        self.title("Enabled Accounts")
        self.overrideredirect(True)
        self.wm_attributes("-topmost", True)
        self.config(bg=self.colors["bg"])
        self.configure(highlightbackground=self.colors["accent"], highlightthickness=1)
        
        w, h = 450, 550
        # Center relative to parent if possible
        try:
            x = parent.winfo_x() + 50
            y = parent.winfo_y() + 50
        except:
             # Fallback
             x, y = 100, 100

        self.geometry("{}x{}+{}+{}".format(w, h, x, y))
        
        # Header
        header = tk.Frame(self, bg=self.colors["bg"], height=40)
        header.pack(fill="x", side="top")
        header.bind("<Button-1>", self.start_move)
        header.bind("<B1-Motion>", self.on_move)
        
        lbl = tk.Label(header, text="Select Accounts", bg=self.colors["bg"], fg=self.colors["fg"], 
                       font=("Segoe UI", 11, "bold"))
        lbl.pack(side="left", padx=15, pady=10)
        
        btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")
        btn_close.pack(side="right", padx=15)
        btn_close.bind("<Button-1>", lambda e: self.destroy())

        # Content (Use reused UI)
        container = tk.Frame(self, bg=self.colors["bg"])
        container.pack(fill="both", expand=True, padx=2, pady=2)
        
        self.ui_helper = AccountSelectionUI(container, accounts, current_enabled, self.launch_folder_selection, bg_color=self.colors["bg"])
        self.ui_helper.pack(fill="both", expand=True)

        # Footer
        footer = tk.Frame(self, bg=self.colors["bg"], height=60)
        footer.pack(fill="x", side="bottom", pady=10)
        
        tk.Button(footer, text="Save Changes", command=self.save_selection,
            bg=self.colors["accent"], fg="black", bd=0, font=("Segoe UI", 10, "bold"), padx=20, pady=5).pack(side="right", padx=15)
        
        tk.Button(footer, text="Cancel", command=self.destroy,
            bg="#333333", fg="white", bd=0, font=("Segoe UI", 10), padx=15, pady=5).pack(side="right", padx=5)

    def save_selection(self):
        final = self.ui_helper.get_settings()
        self.callback(final)
        self.destroy()

    def launch_folder_selection(self, account_name, on_selected, selected_paths=None):
        # Adapted from previous open_folder_picker
        try:
             # Find SidebarWindow reliably
             sidebar = None
             if hasattr(self.master, "outlook_client"):
                 sidebar = self.master
             elif hasattr(self.master, "main_window"):
                 sidebar = self.master.main_window
             elif hasattr(self.master, "master"):
                 sidebar = self.master.master
                 
             if not sidebar or not hasattr(sidebar, "outlook_client"):
                 # messagebox.showerror("Error", "Could not connect to Outlook Sidebar.")
                 print("Error: Could not locate sidebar for folder picker logic")
                 return

             folders = sidebar.outlook_client.get_folder_list(account_name)
             
             if not folders:
                 messagebox.showwarning("No Folders", "Could not retrieve folder list for '{}'.".format(account_name))
                 return
                 
             FolderPickerWindow(self, folders, on_selected, selected_paths, colors=self.colors)
        except Exception as e:
            print("Error opening folder picker: {}".format(e))
            messagebox.showerror("Error", "Failed to open folder picker:\n{}".format(e))


    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def on_move(self, event):
        deltax = event.x - self._x
        deltay = event.y - self._y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry("+{}+{}".format(x, y))
