# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes
import time
import json
import os
import win32com.client
import win32gui
import win32con
import re
import math # Added for animation
import glob
from tkinter import messagebox
from PIL import Image, ImageTk
from datetime import datetime, timedelta

# --- Store Compatibility Imports ---
import sys
import shutil
# Using ctypes for Mutex to avoid extra pywin32 module dependencies if not strictly needed,
# though win32event is also fine since win32gui is used.
# sticking to ctypes kernel32 for zero-dependency bloat for this specific feature.
kernel32 = ctypes.windll.kernel32


# --- Application Constants ---
VERSION = "v1.0.7"


# --- Windows API Constants & Structures ---
ABM_NEW = 0x00000000
ABM_REMOVE = 0x00000001
ABM_QUERYPOS = 0x00000002
ABM_SETPOS = 0x00000003
ABM_GETSTATE = 0x00000004
ABM_GETTASKBARPOS = 0x00000005
ABM_ACTIVATE = 0x00000006
ABM_GETAUTOHIDEBAR = 0x00000007
ABM_SETAUTOHIDEBAR = 0x00000008
ABM_WINDOWPOSCHANGED = 0x00000009
ABM_SETSTATE = 0x0000000A

ABE_LEFT = 0
ABE_TOP = 1
ABE_RIGHT = 2
ABE_BOTTOM = 3

class APPBARDATA(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('hWnd', wintypes.HWND),
        ('uCallbackMessage', wintypes.UINT),
        ('uEdge', wintypes.UINT),
        ('rc', wintypes.RECT),
        ('lParam', wintypes.LPARAM),
    ]

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('rcMonitor', wintypes.RECT),
        ('rcWork', wintypes.RECT),
        ('dwFlags', wintypes.DWORD),
    ]

shell32 = ctypes.windll.shell32
user32 = ctypes.windll.user32

class AppBarManager:
    """
    Manages the Windows AppBar registration and positioning.
    """
    def __init__(self, hwnd, edge=ABE_LEFT):
        self.hwnd = hwnd
        self.edge = edge
        self.registered = False
        self.uCallbackMessage = 0x0400 + 1  # WM_USER + 1
        
        # Keep appbar data persistent
        self.abd = APPBARDATA()
        self.abd.cbSize = ctypes.sizeof(APPBARDATA)
        self.abd.hWnd = self.hwnd
        self.abd.uCallbackMessage = self.uCallbackMessage
        self.abd.uEdge = self.edge

    def register(self):
        """Registers the window as an AppBar."""
        if not self.registered:
            shell32.SHAppBarMessage(ABM_NEW, ctypes.byref(self.abd))
            self.registered = True

    def unregister(self):
        """Unregisters the AppBar."""
        if self.registered:
            shell32.SHAppBarMessage(ABM_REMOVE, ctypes.byref(self.abd))
            self.registered = False

    def set_pos(self, width, monitor_left, monitor_top, monitor_width, monitor_height):
        """
        Reserving space for the AppBar on the current monitor's edge.
        """
        if not self.registered:
            return

        # 1. Query Position
        if self.edge == ABE_LEFT:
            self.abd.rc.left = monitor_left
            self.abd.rc.top = monitor_top
            self.abd.rc.right = monitor_left + width
            self.abd.rc.bottom = monitor_top + monitor_height
        elif self.edge == ABE_RIGHT:
            self.abd.rc.left = monitor_left + monitor_width - width
            self.abd.rc.top = monitor_top
            self.abd.rc.right = monitor_left + monitor_width
            self.abd.rc.bottom = monitor_top + monitor_height
        
        # Query the system for an approved position
        shell32.SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(self.abd))
        
        # 2. Adjust if necessary
        if self.edge == ABE_LEFT:
            self.abd.rc.right = self.abd.rc.left + width
        elif self.edge == ABE_RIGHT:
            self.abd.rc.left = self.abd.rc.right - width
        
        # 3. Set Position
        shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(self.abd))
        
        # Return the actual rectangle committed
        return self.abd.rc.left, self.abd.rc.top, self.abd.rc.right - self.abd.rc.left, self.abd.rc.bottom - self.abd.rc.top

class ScrollableFrame(tk.Frame):
    """
    A scrollable frame that can contain multiple email cards.
    """
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=kwargs.get("bg", "#222222"), highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas, bg=kwargs.get("bg", "#222222"))

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Custom Scroll Buttons
        # We place them relative to 'self' so they overlay the canvas
        self.btn_up = tk.Button(self, text="▲", command=lambda: self.scroll(-1), 
                                bg="#444444", fg="white", bd=0, font=("Arial", 6), width=10, activebackground="#555555", activeforeground="white")
        self.btn_down = tk.Button(self, text="▼", command=lambda: self.scroll(1), 
                                  bg="#444444", fg="white", bd=0, font=("Arial", 6), width=10, activebackground="#555555", activeforeground="white")

        self.canvas.configure(yscrollcommand=self._on_scroll_update)
        self.canvas.pack(side="left", fill="both", expand=True)
        # self.scrollbar removed
        
        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Ensure scrollable frame matches canvas width
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def scroll(self, direction):
        self.canvas.yview_scroll(direction, "units")

    def _on_scroll_update(self, first, last):
        # first and last are strings "0.0" to "1.0"
        f = float(first)
        l = float(last)
        
        if f <= 0.001:
            self.btn_up.place_forget()
        else:
            self.btn_up.place(relx=0.5, rely=0, anchor="n", height=15, relwidth=1.0)
            self.btn_up.lift() # Ensure on top
            
        if l >= 0.999:
            self.btn_down.place_forget()
        else:
            self.btn_down.place(relx=0.5, rely=1.0, anchor="s", height=15, relwidth=1.0)
            self.btn_down.lift() # Ensure on top
        
    def _on_canvas_configure(self, event):
        # Resize the inner frame to match the canvas width
        self.canvas.itemconfig(self.window_id, width=event.width)


class RoundedFrame(tk.Canvas):
    def __init__(self, parent, width, height, corner_radius, padding, color, bg, **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, bg=bg, bd=0, highlightthickness=0, **kwargs)
        self.radius = corner_radius
        self.padding = padding
        self.color = color
        
        self.id = self.create_rounded_rect(0, 0, width, height, self.radius, fill=self.color, outline="")
        
        # Inner frame for widgets
        self.inner = tk.Frame(self, bg=self.color)
        self.window_id = self.create_window((padding, padding), window=self.inner, anchor="nw")
        
        self.bind("<Configure>", self._on_resize)
        
    def _on_resize(self, event):
        self.coords(self.id, self._rounded_rect_coords(0, 0, event.width, event.height, self.radius))
        self.itemconfig(self.window_id, width=event.width - 2*self.padding, height=event.height - 2*self.padding)
        
    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        return self.create_polygon(self._rounded_rect_coords(x1, y1, x2, y2, r), **kwargs)

    def _rounded_rect_coords(self, x1, y1, x2, y2, r):
        points = [x1+r, y1,
                  x1+r, y1,
                  x2-r, y1,
                  x2-r, y1,
                  x2, y1,
                  x2, y1+r,
                  x2, y1+r,
                  x2, y2-r,
                  x2, y2-r,
                  x2, y2,
                  x2-r, y2,
                  x2-r, y2,
                  x1+r, y2,
                  x1+r, y2,
                  x1, y2,
                  x1, y2-r,
                  x1, y2-r,
                  x1, y1+r,
                  x1, y1+r,
                  x1, y1]
        return points

class ToolTip:
    """
    Creates a popup tooltip for a given widget.
    """
    def __init__(self, widget, text, side="bottom"):
        self.widget = widget
        self.text = text
        self.side = side # "bottom", "left", "right", "top"
        self.tip_window = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.show_tip()

    def leave(self, event=None):
        self.hide_tip()

    def show_tip(self):
        """Displays the tooltip."""
        if self.tip_window or not self.text:
            return
        
        # Create window first to get size
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        
        label = tk.Label(
            tw, 
            text=self.text, 
            justify="left",
            bg="#2d2d2d", 
            fg="#ffffff",
            relief="solid", 
            borderwidth=1,
            font=("Segoe UI", 8),
            padx=4, pady=2
        )
        label.pack(ipadx=1)
        
        tw.update_idletasks() # Calculate size
        
        tw_width = tw.winfo_reqwidth()
        tw_height = tw.winfo_reqheight()
        
        widget_x = self.widget.winfo_rootx()
        widget_y = self.widget.winfo_rooty()
        widget_w = self.widget.winfo_width()
        widget_h = self.widget.winfo_height()
        
        if self.side == "left":
            x = widget_x - tw_width - 5
            y = widget_y + (widget_h // 2) - (tw_height // 2)
        elif self.side == "right":
            x = widget_x + widget_w + 5
            y = widget_y + (widget_h // 2) - (tw_height // 2)
        elif self.side == "top":
            x = widget_x + (widget_w // 2) - (tw_width // 2)
            y = widget_y - tw_height - 5
        else: # bottom
            x = widget_x + 20
            y = widget_y + widget_h + 5
            
        tw.wm_geometry(f"+{x}+{y}")

    def hide_tip(self):
        """Hides the tooltip."""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class OutlookClient:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.last_received_time = None
        self.connect()
        # Initialize last_received_time
        if self.namespace:
            self.check_latest_time()

    def connect(self):
        """Attempts to connect to the Outlook COM object."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            # print("Connected to Outlook")
            return True
        except Exception as e:
            print(f"Error connecting to Outlook: {e}")
            self.outlook = None
            self.namespace = None
            return False

    def check_latest_time(self):
        """Initializes or updates the last received time without returning bool."""
        if not self.namespace: return
        try:
            inbox = self.namespace.GetDefaultFolder(6)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            item = items.GetFirst()
            if item:
                self.last_received_time = item.ReceivedTime
        except Exception:
             # If we fail here, we might be disconnected, but this is just init
             pass

    def check_new_mail(self):
        """Checks if there is email newer than the last check. Recovers connection if needed."""
        # Retry loop (Try once, if fail, reconnect and try again)
        for attempt in range(2):
            if not self.namespace:
                if not self.connect():
                    return False # Still cannot connect

            try:
                inbox = self.namespace.GetDefaultFolder(6)
                items = inbox.Items
                items.Sort("[ReceivedTime]", True)
                item = items.GetFirst()
                
                if item:
                    current_time = item.ReceivedTime
                    # If we have a stored time and the new one is newer
                    if self.last_received_time and current_time > self.last_received_time:
                        self.last_received_time = current_time
                        return True
                    
                    # Update tracker regardless to avoid stale alerts
                    self.last_received_time = current_time
                return False # No new mail
                
            except Exception as e:
                print(f"Polling error (Attempt {attempt+1}): {e}")
                self.namespace = None # Force reconnect next loop
        
        return False

    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, include_read_flagged=True, flag_date_filter="Anytime"):
        # Retry loop
        for attempt in range(2):
            if not self.namespace:
                if not self.connect():
                    return []

            try:
                inbox = self.namespace.GetDefaultFolder(6) # 6 = olFolderInbox
                items = inbox.Items
                
                # Build restriction string
                restricts = []
                
                if only_flagged:
                    # [FlagStatus] <> 0 correctly identifies flagged items
                    restricts.append("[FlagStatus] <> 0")
                    # Sub-filter: Only apply unread restriction if NOT including read flagged
                    if not include_read_flagged:
                        restricts.append("[UnRead] = True")
                    
                    # Date Filter Logic
                    if flag_date_filter and flag_date_filter != "Anytime":
                        now = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                        if flag_date_filter == "Today":
                            start = now
                            end = now + timedelta(days=1)
                            restricts.append(f"[TaskDueDate] >= '{start.strftime('%m/%d/%Y %I:%M %p')}' AND [TaskDueDate] < '{end.strftime('%m/%d/%Y %I:%M %p')}'")
                        elif flag_date_filter == "Tomorrow":
                            start = now + timedelta(days=1)
                            end = now + timedelta(days=2)
                            restricts.append(f"[TaskDueDate] >= '{start.strftime('%m/%d/%Y %I:%M %p')}' AND [TaskDueDate] < '{end.strftime('%m/%d/%Y %I:%M %p')}'")
                        elif flag_date_filter == "This Week":
                            # Next 7 days
                            end = now + timedelta(days=7)
                            restricts.append(f"[TaskDueDate] >= '{now.strftime('%m/%d/%Y %I:%M %p')}' AND [TaskDueDate] < '{end.strftime('%m/%d/%Y %I:%M %p')}'")
                        elif flag_date_filter == "Next Week":
                            # Days 7 to 14
                            start = now + timedelta(days=7)
                            end = now + timedelta(days=14)
                            restricts.append(f"[TaskDueDate] >= '{start.strftime('%m/%d/%Y %I:%M %p')}' AND [TaskDueDate] < '{end.strftime('%m/%d/%Y %I:%M %p')}'")
                        elif flag_date_filter == "No Date":
                            # Outlook uses 4501-01-01 for "No Date" tasks
                            restricts.append("[TaskDueDate] > '01/01/4500'")
                else:
                    # Global filter: Only unread
                    if unread_only:
                        restricts.append("[UnRead] = True")
                
                if restricts:
                    restrict_str = " AND ".join(restricts)
                    items = items.Restrict(restrict_str)
                
                items.Sort("[ReceivedTime]", True) # Descending
                
                email_list = []
                for i, item in enumerate(items):
                    if i >= count:
                        break
                    try:
                        subject = getattr(item, "Subject", "[No Subject]")
                        sender = getattr(item, "SenderName", "Unknown")
                        raw_body = getattr(item, "Body", "")
                        
                        # Clean up body: remove newlines and extra spaces
                        clean_body = re.sub(r'\s+', ' ', raw_body).strip()
                        body = clean_body[:100] + "..." # Preview
                        
                        unread = getattr(item, "UnRead", False)
                        
                        email_list.append({
                            "sender": sender,
                            "subject": subject,
                            "preview": body,
                            "unread": unread,
                            "entry_id": getattr(item, "EntryID", ""),
                            "received": getattr(item, "ReceivedTime", None),
                            "flag_status": getattr(item, "FlagStatus", 0),
                            "due_date": getattr(item, "TaskDueDate", None)
                        })
                    except Exception as inner_e:
                        print(f"Error reading item: {inner_e}")
                        
                return email_list
            except Exception as e:
                print(f"Fetch error (Attempt {attempt+1}): {e}")
                self.namespace = None # Force reconnect
        
        
        return []

    def get_item_by_entryid(self, entry_id):
        """Retrieves a specific Outlook item by its EntryID."""
        if not self.namespace:
            self.connect()
        try:
            return self.namespace.GetItemFromID(entry_id)
        except Exception as e:
            print(f"Error getting item {entry_id}: {e}")
            return None

    def find_folder_by_name(self, folder_name):
        """
        Recursively searches for a folder by name. 
        Starts at default Inbox parent (likely the account root).
        This is a simple implementation; heavy trees might need optimization.
        """
        if not self.namespace: return None
        
        try:
            # Start at root of the default store
            root = self.namespace.GetDefaultFolder(6).Parent # Inbox -> Parent (Account Root)
            
            # Helper for recursion
            def recursive_find(folder):
                if folder.Name.lower() == folder_name.lower():
                    return folder
                for sub in folder.Folders:
                    found = recursive_find(sub)
                    if found: return found
                return None

            return recursive_find(root)
        except Exception as e:
            print(f"Error finding folder {folder_name}: {e}")
            return None

    def get_folder_list(self):
        """Returns a list of folder paths (e.g. 'Inbox', 'Inbox/ProjectA')"""
        if not self.namespace: return []
        
        folders = []
        try:
            root = self.namespace.GetDefaultFolder(6).Parent
            
            def recurse(folder, parent_path=""):
                try:
                    name = folder.Name
                    path = f"{parent_path}/{name}" if parent_path else name
                    
                    # Add to list
                    folders.append(path)
                    
                    # Limit recursion depth to avoid slowdowns on massive mailboxes
                    # Only go 2 levels deep for now? Or just try all.
                    # Let's do 1 level deep for safety in this version.
                    if parent_path.count("/") < 2:
                        for sub in folder.Folders:
                            recurse(sub, path)
                except Exception:
                    pass

            # Start recursion
            for f in root.Folders:
                recurse(f)
                
        except Exception as e:
            print(f"Error fetching folder list: {e}")
            
        return sorted(folders)

class FolderPickerWindow(tk.Toplevel):
    def __init__(self, parent, folders, callback):
        super().__init__(parent)
        self.callback = callback
        self.folders = folders
        
        # Win11 Colors
        self.colors = {
            "bg": "#202020",
            "fg": "#FFFFFF",
            "accent": "#60CDFF", 
            "select_bg": "#444444"
        }
        
        self.overrideredirect(True)
        self.wm_attributes("-topmost", True)
        self.config(bg=self.colors["bg"])
        self.configure(highlightbackground=self.colors["accent"], highlightthickness=1)
        
        # Geometry
        w, h = 300, 400
        x = parent.winfo_x() + 50
        y = parent.winfo_y() + 50
        self.geometry(f"{w}x{h}+{x}+{y}")

        # Title Bar
        header = tk.Frame(self, bg=self.colors["bg"], height=30)
        header.pack(fill="x", side="top")
        header.bind("<Button-1>", self.start_move)
        header.bind("<B1-Motion>", self.on_move)

        lbl = tk.Label(header, text="Select Folder", bg=self.colors["bg"], fg=self.colors["fg"], font=("Segoe UI", 10, "bold"))
        lbl.pack(side="left", padx=10, pady=5)
        
        btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")
        btn_close.pack(side="right", padx=10)
        btn_close.bind("<Button-1>", lambda e: self.destroy())

        # TreeView
        tree_frame = tk.Frame(self, bg=self.colors["bg"])
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
            background="#2D2D30", 
            foreground="white", 
            fieldbackground="#2D2D30",
            borderwidth=0
        )
        style.map("Treeview", background=[("selected", self.colors["accent"])])

        self.tree = ttk.Treeview(tree_frame, show="tree", selectmode="browse")
        self.tree.pack(side="left", fill="both", expand=True)
        
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        
        self.populate_tree()
        
        # Select Button
        btn_sel = tk.Button(self, text="Select", command=self.select_folder,
            bg=self.colors["accent"], fg="black", bd=0, font=("Segoe UI", 9, "bold"), pady=5)
        btn_sel.pack(fill="x", padx=10, pady=10)

    def populate_tree(self):
        # Build hierarchy from paths e.g. "Inbox/Sub"
        # Since we just have paths, we can just list them flat or try to build structure.
        # User asked for mirror of Outlook. Start simple: Tree nodes.
        
        # Logic to build tree from slash-paths
        nodes = {}
        
        for path in self.folders:
            parts = path.split("/")
            parent = ""
            current = ""
            
            for i, part in enumerate(parts):
                current = f"{parent}/{part}" if parent else part
                
                # Check if node exists
                if current not in nodes:
                    # Parent ID is clean parent path logic
                    pid = parent if parent else ""
                    
                    # Insert
                    nodes[current] = self.tree.insert(pid, "end", iid=current, text=part, open=False)
                
                parent = current

    def select_folder(self):
        sel = self.tree.selection()
        if sel:
            # The IID is the full path in our logic
            path = sel[0]
            self.callback(path)
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


class SettingsPanel(tk.Frame):
    """Inline settings panel that extends from the sidebar."""
    def __init__(self, parent, main_window, callback):
        super().__init__(parent, bg="#202020")
        self.main_window = main_window
        self.callback = callback
        
        # --- Windows 11 Dark Theme ---
        self.colors = {
            "bg_root": "#202020",       # Deep Dark
            "bg_card": "#2D2D30",       # Input BG
            "accent": "#60CDFF",        # Win11 Blue
            "fg_text": "#FFFFFF",
            "fg_dim": "#A0A0A0",
            "input_bg": "#333333"       # Distinct slot color
        }
        
        # Configure ttk Theme
        style = ttk.Style(self)
        style.theme_use("clam")
        
        # TCombobox - Flat, Dark
        style.configure("TCombobox", 
            fieldbackground=self.colors["input_bg"], 
            background=self.colors["bg_card"], 
            foreground="white",
            arrowcolor="white",
            bordercolor=self.colors["bg_root"],
            darkcolor=self.colors["bg_root"],
            lightcolor=self.colors["bg_root"]
        )
        style.map("TCombobox", fieldbackground=[("readonly", self.colors["input_bg"])])
        
        # TEntry - Flat, Dark
        style.configure("TEntry", 
            fieldbackground=self.colors["input_bg"], 
            foreground="white",
            bordercolor=self.colors["bg_root"],
            lightcolor=self.colors["bg_root"],
            darkcolor=self.colors["bg_root"]
        )
        
        # Frame styling
        self.config(bg=self.colors["bg_root"])
        self.configure(highlightbackground="#444444", highlightthickness=1)
        
        # Fixed width for the settings panel
        self.panel_width = 370
        self.config(width=self.panel_width)
        self.pack_propagate(False)  # Prevent shrinking
        
        # --- Header ---
        header = tk.Frame(self, bg=self.colors["bg_root"], height=40)
        header.pack(fill="x", side="top")
        
        lbl_title = tk.Label(header, text="Settings", fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Segoe UI Variable Display", 12, "bold"))
        lbl_title.pack(side="left", padx=20, pady=10)
        
        title_underline = tk.Frame(self, bg=self.colors["accent"], height=2)
        title_underline.pack(fill="x", side="top")

        # Configure larger font for dropdown lists (affects all comboboxes in this window ideally, but mainly for icons)
        self.option_add('*TCombobox*Listbox.font', ("Segoe UI", 16))

        # Red Cross Close
        btn_close = tk.Label(header, text="✕", fg="#FFFFFF", bg="#C42B1C", font=("Arial", 10), width=5, cursor="hand2")
        btn_close.pack(side="right", fill="y", padx=0)
        btn_close.bind("<Button-1>", lambda e: self.close_panel())

        # Attribution Info Button
        btn_info = tk.Label(header, text="ⓘ", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 12), cursor="hand2")
        btn_info.pack(side="right", padx=10)
        ToolTip(btn_info, "Icons made by IconKanan and Ardiansyah from www.flaticon.com", side="left")
        
        # Helper method to create section headers
        def create_section_header(parent, title):
            """Creates a section header with title and divider line."""
            section_frame = tk.Frame(parent, bg=self.colors["bg_root"])
            section_frame.pack(fill="x", padx=2, pady=(15, 5))
            
            # Title label
            tk.Label(section_frame, text=title, bg=self.colors["bg_root"], fg=self.colors["fg_text"], 
                    font=("Segoe UI", 11, "bold")).pack(side="left", anchor="w")
            
            # Divider line (partial width)
            divider = tk.Frame(section_frame, bg="#555555", height=1)
            divider.pack(side="left", fill="x", expand=True, padx=(10, 0))
        
        # === SECTION 1: Button Selection ===
        create_section_header(self, "Button Selection")
        
        # --- Button Configuration Table ---
        container = tk.Frame(self, bg=self.colors["bg_root"], pady=20)
        container.pack(fill="x", expand=False, padx=(2, 20))  # 2px left padding
        
        # Table Headers
        headers = ["Icon", "Action", "Folders (for Move)"]
        
        for col, text in enumerate(headers):
            tk.Label(
                container, text=text, 
                bg=self.colors["bg_root"], fg=self.colors["fg_dim"], 
                font=("Segoe UI", 9)
            ).grid(row=0, column=col, sticky="w", padx=8, pady=(0, 15))
            
        # Rows
        self.rows_data = [] 
        self.action_options = ["None", "Mark Read", "Delete", "Read & Delete", "Flag", "Open Email", "Reply", "Move To..."]
        # Monochrome / Clean Unicode Icons AND Custom PNGs
        unicode_icons = ["", "🗑", "✉", "⚑", "↩", "📂", "↗", "✓", "✕", "⚠"]
        
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
            
            # 1. Icon - Replaced with Spacer as per user request (Space reserved)
            # cb_icon = ttk.Combobox(container, values=self.icons, width=3, state="readonly", font=("Segoe UI", 16))
            # cb_icon.set(c_data.get("icon", self.icons[0]))
            # cb_icon.grid(row=i+1, column=0, padx=8, pady=8, ipady=4)
            
            # 1. Icon Display (Dynamic Label)
            lbl_icon = tk.Label(container, bg=self.colors["bg_root"], width=5) # Width roughly matches 30px
            lbl_icon.grid(row=i+1, column=0, padx=8, pady=8)
            
            # Preserve the icon value for saving (start with current)
            current_icon_val = c_data.get("icon", self.icons[0])
            
            # 2. Action (Previously Action 1)
            cb_act1 = ttk.Combobox(container, values=self.action_options, width=15, state="readonly", font=("Segoe UI", 10))
            cb_act1.set(c_data.get("action1", "None")) 
            cb_act1.grid(row=i+1, column=1, padx=8, pady=8, ipady=4)
            
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
                             img = self.main_window.load_icon_white(path, size=(24, 24))
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
                
                # Update underlying data for saving
                # We need to update the entry in self.rows_data list, 
                # but we are currently building it. 
                # Better approach: Modify the specific dictionary in rows_data via mutable access
                # BUT rows_data isn't fully populated yet.
                # So we bind a function that looks up the row data LATER.
                if len(self.rows_data) > row_idx:
                    self.rows_data[row_idx]["icon_val"] = new_icon

            # Initial Update
            # We use a deferred call or just run it now contextually, but we need the 'row_idx' 
            # which is loop variable 'i'. Be careful with closures.
            
            # Auto-update Handler
            def on_action_change(event, act_cb=cb_act1, icon_lbl=lbl_icon, idx=i):
                 update_icon_display(act_cb, icon_lbl, idx)
                 self.refresh_dropdown_options() # Enforce uniqueness
                 self.update_button_config()  # Apply changes immediately
            
            cb_act1.bind("<<ComboboxSelected>>", on_action_change)
            
            # 3. Folder Picker UI (Entry + Button) - Shifted to Column 2
            f_frame = tk.Frame(container, bg=self.colors["bg_root"])
            f_frame.grid(row=i+1, column=2, padx=8, pady=8)
            
            e_folder = ttk.Entry(f_frame, width=15, font=("Segoe UI", 10))
            e_folder.insert(0, c_data.get("folder", ""))
            e_folder.pack(side="left", ipady=4)

            # Picker Button
            btn_pick = tk.Label(f_frame, text="...", bg=self.colors["bg_card"], fg="white", font=("Segoe UI", 8), width=3, cursor="hand2")
            btn_pick.pack(side="left", padx=(5,0), fill="y")
            
            # Bind picker
            def open_picker(event, entry=e_folder):
                folders = self.main_window.outlook_client.get_folder_list()
                FolderPickerWindow(self, folders if folders else ["Inbox"], 
                                   lambda path: (entry.delete(0, tk.END), entry.insert(0, path)))

            btn_pick.bind("<Button-1>", open_picker)

            self.rows_data.append({
                "icon_val": current_icon_val, # Store value directly
                "act1": cb_act1,
                "folder": e_folder
            })
            
            # Trigger initial display update manually
            update_icon_display(cb_act1, lbl_icon, i)
            
        # Initial Refresh of Options to filter out duplicates
        self.refresh_dropdown_options()
            
        # --- Sidebar Placement Setting REMOVED (Auto-snap implemented) ---
        # placement_frame = tk.Frame(self, bg=self.colors["bg_root"])
        # ...

        # === SECTION 2: General Settings ===
        create_section_header(self, "General Settings")

        # --- Typography Setting ---
        typo_frame = tk.Frame(self, bg=self.colors["bg_root"])
        typo_frame.pack(fill="x", padx=(20, 30), pady=(10, 0))
        
        tk.Label(typo_frame, text="Font Family:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        self.font_fam_cb = ttk.Combobox(typo_frame, values=["Segoe UI", "Arial", "Verdana", "Tahoma", "Courier New", "Georgia"], width=15, state="readonly", font=("Segoe UI", 12))
        self.font_fam_cb.set(self.main_window.font_family)
        self.font_fam_cb.pack(side="left", padx=(5, 20))
        self.font_fam_cb.bind("<<ComboboxSelected>>", self.update_font_settings)
        
        tk.Label(typo_frame, text="Size:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        self.font_size_cb = ttk.Combobox(typo_frame, values=[str(i) for i in range(8, 17)], width=5, state="readonly", font=("Segoe UI", 12))
        self.font_size_cb.set(str(self.main_window.font_size))
        self.font_size_cb.pack(side="left", padx=5)
        self.font_size_cb.bind("<<ComboboxSelected>>", self.update_font_settings)
        
        # --- System Settings (Refresh Rate) ---
        self.refresh_options = {"15s": 15, "30s": 30, "1m": 60, "2m": 120, "5m": 300}
        sys_frame = tk.Frame(self, bg=self.colors["bg_root"])
        sys_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        tk.Label(sys_frame, text="Refresh Rate:", fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 10)).pack(side="left")
        self.refresh_cb = ttk.Combobox(sys_frame, values=list(self.refresh_options.keys()), width=10, state="readonly", font=("Segoe UI", 12))
        
        current_label = "30s"
        for label, val in self.refresh_options.items():
            if val == self.main_window.poll_interval:
                current_label = label
                break
        self.refresh_cb.set(current_label)
        self.refresh_cb.pack(side="left", padx=5)
        self.refresh_cb.bind("<<ComboboxSelected>>", self.update_refresh_rate)

        # === SECTION 3: Window Selection ===
        create_section_header(self, "Window Selection")
        
        # --- Window Mode Selector ---
        window_frame = tk.Frame(self, bg=self.colors["bg_root"])
        window_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        # Track window mode (single or dual)
        self.window_mode_var = tk.StringVar(value="single")  # Default to single
        
        # Single Window Button
        self.btn_single_window = tk.Button(
            window_frame, text="Single Window", 
            command=lambda: self.select_window_mode("single"),
            bg=self.colors["accent"], fg="black",
            font=("Segoe UI", 10, "bold"),
            bd=0, padx=20, pady=8,
            activebackground=self.colors["accent"],
            activeforeground="black"
        )
        self.btn_single_window.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        # Dual Window Button
        self.btn_dual_window = tk.Button(
            window_frame, text="Dual Window", 
            command=lambda: self.select_window_mode("dual"),
            bg=self.colors["bg_card"], fg="white",
            font=("Segoe UI", 10),
            bd=0, padx=20, pady=8,
            activebackground=self.colors["bg_card"],
            activeforeground="white"
        )
        self.btn_dual_window.pack(side="left", fill="x", expand=True)

        # === SECTION 4: Email Settings ===
        create_section_header(self, "Email Settings")

        # --- Email List Settings ---
        list_settings_frame = tk.Frame(self, bg=self.colors["bg_root"])
        list_settings_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        self.show_read_var = tk.BooleanVar(value=self.main_window.show_read)
        self.chk_show_read = tk.Checkbutton(
            list_settings_frame, text="Include read email", 
            variable=self.show_read_var,
            command=self.update_email_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 10)
        )
        self.chk_show_read.grid(row=0, column=0, sticky="w", pady=(0, 5))

        # === SECTION 5: Reminder Settings ===
        create_section_header(self, "Reminder Settings")
        
        # --- Flagged Email Filters ---
        flagged_settings_frame = tk.Frame(self, bg=self.colors["bg_root"])
        flagged_settings_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))

        # Flagged Filter Row
        self.only_flagged_var = tk.BooleanVar(value=self.main_window.only_flagged)
        
        # Define update_flag_options here so it can be used in command
        def update_flag_options(event=None):
            state = "normal" if self.only_flagged_var.get() else "disabled"
            self.chk_include_read_flagged.config(state=state)
            self.flag_date_cb.config(state=state)
            self.update_email_filters()  # Apply changes

        self.chk_only_flagged = tk.Checkbutton(
            flagged_settings_frame, text="Only Flagged Emails", 
            variable=self.only_flagged_var,
            command=update_flag_options,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 10)
        )
        self.chk_only_flagged.grid(row=0, column=0, sticky="w")

        self.include_read_flagged_var = tk.BooleanVar(value=self.main_window.include_read_flagged)
        self.chk_include_read_flagged = tk.Checkbutton(
            flagged_settings_frame, text="Include Read", 
            variable=self.include_read_flagged_var,
            command=self.update_email_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 10)
        )
        self.chk_include_read_flagged.grid(row=1, column=0, sticky="w", padx=(20, 0))

        # Date Filter for Flagged
        tk.Label(flagged_settings_frame, text="Due:", bg=self.colors["bg_root"], fg=self.colors["fg_dim"], font=("Segoe UI", 10)).grid(row=1, column=1, sticky="w", padx=(20, 5))
        self.flag_date_cb = ttk.Combobox(
            flagged_settings_frame, 
            values=["Anytime", "Today", "Tomorrow", "This Week", "Next Week", "No Date"],
            width=10, state="readonly", font=("Segoe UI", 10)
        )
        self.flag_date_cb.set(self.main_window.flag_date_filter)
        self.flag_date_cb.grid(row=1, column=2, sticky="w")
        self.flag_date_cb.bind("<<ComboboxSelected>>", self.update_email_filters)

        # Disable sub-options if No Flagged Filter
        
        update_flag_options() # Initial state

        # --- Icon Brightness Setting REMOVED ---
        # User requested fixed 75% brightness, slider removed.
        
        # Version Label
        lbl_ver = tk.Label(self, text=VERSION, fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 8))
        lbl_ver.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)

    def refresh_dropdown_options(self):
        """Filters available options for each dropdown to prevent duplicate selections."""
        # 1. Collect all currently selected actions (exclude "None")
        selected_actions = []
        for row in self.rows_data:
            act = row["act1"].get()
            if act and act != "None":
                selected_actions.append(act)
        
        # 2. Update each dropdown
        for row in self.rows_data:
            cb = row["act1"]
            current = cb.get()
            
            # Allowed = All Options - (Selected Actions - My Selection)
            # Basically, everything is allowed EXCEPT what others have picked.
            # My current selection must remain valid/in list.
            
            unavailable = [x for x in selected_actions if x != current]
            new_values = [opt for opt in self.action_options if opt not in unavailable]
            
            cb.config(values=new_values)

    def update_font_settings(self, event=None):
        """Apply font changes immediately."""
        self.main_window.font_family = self.font_fam_cb.get()
        try:
            self.main_window.font_size = int(self.font_size_cb.get())
        except:
            self.main_window.font_size = 9
        self.main_window.save_config()
        self.callback()  # refresh_emails

    def update_refresh_rate(self, event=None):
        """Apply refresh rate change immediately."""
        self.main_window.poll_interval = self.refresh_options.get(self.refresh_cb.get(), 30)
        self.main_window.save_config()

    def update_email_filters(self, event=None):
        """Apply email filter changes immediately."""
        self.main_window.show_read = self.show_read_var.get()
        self.main_window.only_flagged = self.only_flagged_var.get()
        self.main_window.include_read_flagged = self.include_read_flagged_var.get()
        self.main_window.flag_date_filter = self.flag_date_cb.get()
        self.main_window.save_config()
        self.callback()  # refresh_emails

    def update_button_config(self):
        """Apply button config changes immediately."""
        new_config = []
        count = 0
        for data in self.rows_data:
            act1 = data["act1"].get()
            if act1 != "None":
                count += 1
                new_config.append({
                    "icon": data["icon_val"],
                    "action1": act1,
                    "folder": data["folder"].get()
                })
        
        self.main_window.btn_count = count
        self.main_window.btn_config = new_config
        self.main_window.save_config()
        self.callback()  # refresh_emails

    def select_window_mode(self, mode):
        """Handle window mode selection (single or dual)."""
        self.window_mode_var.set(mode)
        
        # Update button styling to show active state
        if mode == "single":
            # Single is active
            self.btn_single_window.config(
                bg=self.colors["accent"], 
                fg="black",
                font=("Segoe UI", 10, "bold")
            )
            # Dual is inactive
            self.btn_dual_window.config(
                bg=self.colors["bg_card"], 
                fg="white",
                font=("Segoe UI", 10)
            )
        else:  # dual
            # Dual is active
            self.btn_dual_window.config(
                bg=self.colors["accent"], 
                fg="black",
                font=("Segoe UI", 10, "bold")
            )
            # Single is inactive
            self.btn_single_window.config(
                bg=self.colors["bg_card"], 
                fg="white",
                font=("Segoe UI", 10)
            )
        
        # TODO: Apply window mode to main window layout
        # self.main_window.window_mode = mode
        # self.main_window.save_config()
        # self.main_window.apply_window_layout()

    def close_panel(self):
        """Close the settings panel."""
        self.main_window.toggle_settings_panel()

class SidebarWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- Configuration ---
        self.min_width = 300  
        self.hot_strip_width = 10
        self.expanded_width = 300
        self.is_pinned = True
        self.is_expanded = False
        self.dock_side = "Left" # "Left" or "Right"
        self.font_family = "Segoe UI"
        self.font_size = 9
        self.poll_interval = 30 # seconds
        self.show_read = False
        self.only_flagged = False
        self.include_read_flagged = True
        self.flag_date_filter = "Anytime"
        self.hover_delay = 500 # ms
        self._hover_timer = None
        self._collapse_timer = None
        
        # Settings Panel State
        self.settings_panel_open = False
        self.settings_panel = None
        self.settings_panel_width = 370
        
        # Pulse Animation State
        self.pulsing = False
        self.pulse_step = 0
        self._pulse_job = None
        self.animation_speed = 0.05 # Increment per frame
        self.base_color = "#007ACC"
        self.pulse_color = "#99D9EA" # Lighter cyan/blue for the bar
        
        # Custom Buttons State
        self.btn_count = 2
        # Default config structure update
        self.btn_config = [
            {"label": "Trash", "icon": "✕", "action1": "Mark Read", "action2": "Delete", "folder": ""}, 
            {"label": "Reply", "icon": "↩", "action1": "Reply", "action2": "None", "folder": ""}
        ]
        
        # Load Config
        self.load_config()

        # --- Outlook Client ---
        self.outlook_client = OutlookClient()
        
        # Image Cache (to keep references alive)
        self.image_cache = {}

        # --- Window Setup ---
        self.overrideredirect(True)  # Frameless
        self.wm_attributes("-topmost", True)
        self.config(bg="#333333")

        # Get Screen Dimensions (will be updated in apply_state)
        self.monitor_x = 0
        self.monitor_y = 0
        self.screen_width = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()

        # --- AppBar Manager ---
        self.update_idletasks() 
        self.hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
        if not self.hwnd:
             self.hwnd = self.winfo_id()

        self.appbar = AppBarManager(self.hwnd)
        
        # --- UI Components ---
        # Container frame that holds main content and settings panel side by side
        self.content_wrapper = tk.Frame(self, bg="#222222")
        self.content_wrapper.pack(fill="both", expand=True)
        
        # Main sidebar content frame (expands to fill space when settings closed)
        self.main_frame = tk.Frame(self.content_wrapper, bg="#222222")
        self.main_frame.pack(side="left", fill="both", expand=True)

        # Footer
        self.footer = tk.Frame(self.main_frame, bg="#444444", height=40)
        self.footer.pack(fill="x", side="bottom")
        
        # Footer Buttons
        # Pack order: Rightmost first.
        
        # 1. Outlook Button (Rightmost)
        if os.path.exists("icons/Outlook_48x48.png"):
             # Increase size to 32x32
             try:
                pil_img = Image.open("icons/Outlook_48x48.png").convert("RGBA")
                pil_img = pil_img.resize((32, 32), Image.Resampling.LANCZOS)
                img = ImageTk.PhotoImage(pil_img)
                self.image_cache["outlook_footer"] = img
                self.btn_outlook = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                self.btn_outlook.pack(side="right", padx=(5, 10), pady=5)
                self.btn_outlook.bind("<Button-1>", lambda e: self.open_outlook_app())
                ToolTip(self.btn_outlook, "Open Outlook")
             except Exception as e:
                print(f"Error loading Outlook icon: {e}")

        # 0. Close Button (Leftmost)
        # Use a simple text button or icon if available
        self.btn_close = tk.Label(self.footer, text="✕", bg="#444444", fg="#aaaaaa", font=("Arial", 12), cursor="hand2")
        self.btn_close.pack(side="left", padx=10, pady=5)
        self.btn_close.bind("<Button-1>", lambda e: self.quit_application())
        ToolTip(self.btn_close, "Close Application")
        
        # Version Label
        self.lbl_version = tk.Label(self.footer, text=VERSION, bg="#444444", fg="#888888", font=("Segoe UI", 8))
        self.lbl_version.pack(side="left", padx=5, pady=5)
                 
        # 2. Calendar Button (Next to Outlook)
        if os.path.exists("icons/OutlookCalendar_48x48.png"):
             # Increase size to 32x32
             try:
                pil_img = Image.open("icons/OutlookCalendar_48x48.png").convert("RGBA")
                pil_img = pil_img.resize((32, 32), Image.Resampling.LANCZOS)
                img = ImageTk.PhotoImage(pil_img)
                self.image_cache["calendar_footer"] = img
                self.btn_calendar = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                self.btn_calendar.pack(side="right", padx=5, pady=5)
                self.btn_calendar.bind("<Button-1>", lambda e: self.open_calendar_app())
                ToolTip(self.btn_calendar, "Open Calendar")
             except Exception as e:
                print(f"Error loading Calendar icon: {e}")

        # Header
        self.header = tk.Frame(self.main_frame, bg="#444444", height=40)
        self.header.pack(fill="x", side="top")
        
        # Header Dragging
        self.header.bind("<Button-1>", self.start_window_drag)
        self.header.bind("<B1-Motion>", self.on_window_drag)
        self.header.bind("<ButtonRelease-1>", self.stop_window_drag)
        
        # Title
        self.lbl_title = tk.Label(self.header, text="Outlook Monitor", bg="#444444", fg="white", font=(self.font_family, 10, "bold"))
        self.lbl_title.pack(side="left", padx=10)
        self.lbl_title.bind("<Button-1>", self.start_window_drag)
        self.lbl_title.bind("<B1-Motion>", self.on_window_drag)
        self.lbl_title.bind("<ButtonRelease-1>", self.stop_window_drag)

        # Pin Button / Logo (Custom Canvas)
        self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg="#444444", highlightthickness=0)
        self.btn_pin.pack(side="right", padx=5, pady=5)
        self.btn_pin.bind("<Button-1>", lambda e: self.toggle_pin())
        self.draw_pin_icon()
        
        # Custom Settings Button (Cog)
        if os.path.exists("icons/Settings.png"):
            img = self.load_icon_white("icons/Settings.png", size=(24, 24))
            if img:
                self.image_cache["settings_header"] = img
                self.btn_settings = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
            else:
                 self.btn_settings = tk.Label(self.header, text="⚙", bg="#444444", fg="#aaaaaa", font=(self.font_family, 12), cursor="hand2")
        else:
            self.btn_settings = tk.Label(self.header, text="⚙", bg="#444444", fg="#aaaaaa", font=(self.font_family, 12), cursor="hand2")
        self.btn_settings.pack(side="right", padx=5)
        self.btn_settings.bind("<Button-1>", lambda e: self.open_settings())

        # Refresh Button
        if os.path.exists("icons/Sync.png"):
            # Increased by another 10% -> (28, 28)
            img = self.load_icon_white("icons/Sync.png", size=(28, 28))
            if img:
                self.image_cache["sync_header"] = img
                self.btn_refresh = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
            else:
                 self.btn_refresh = tk.Label(self.header, text="↻", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        else:
            self.btn_refresh = tk.Label(self.header, text="↻", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        self.btn_refresh.pack(side="right", padx=5)
        self.btn_refresh.bind("<Button-1>", lambda e: self.refresh_emails())
        
        ToolTip(self.btn_settings, "Settings")
        ToolTip(self.btn_refresh, "Refresh Email List")

        # Share Button
        if os.path.exists("icons/Share.png"):
            # Reduced by another 10% -> (20, 20)
            img = self.load_icon_white("icons/Share.png", size=(20, 20))
            if img:
                self.image_cache["share_header"] = img
                self.btn_share = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
            else:
                 self.btn_share = tk.Label(self.header, text="🔗", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        else:
            self.btn_share = tk.Label(self.header, text="🔗", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        self.btn_share.pack(side="right", padx=5)
        # No action yet, just tooltip
        ToolTip(self.btn_share, "Sharing not available yet")

        # Content Area - Using grid for precise height control
        # Create container for grid layout (to avoid mixing pack/grid on main_frame)
        grid_container = tk.Frame(self.main_frame, bg="#000000")
        grid_container.pack(expand=True, fill="both", padx=5, pady=5)
        
        # Configure grid rows: row 0 = email (weight 1), row 1 = reminder (weight 1)
        grid_container.rowconfigure(0, weight=1)  # Email row
        grid_container.rowconfigure(1, weight=1)  # Reminder row
        grid_container.columnconfigure(0, weight=1)
        
        # Email container (row 0, 50% height)
        self.content_container = tk.Frame(grid_container, bg="#222222")
        self.content_container.grid(row=0, column=0, sticky="nsew", padx=0, pady=(0, 2))
        
        # Email section header
        email_header = tk.Frame(self.content_container, bg="#333333", height=20)
        email_header.pack(fill="x", side="top")
        email_header.pack_propagate(False)  # Maintain fixed height
        
        tk.Label(
            email_header, text="Email", 
            bg="#333333", fg="#AAAAAA",
            font=(self.font_family, 9, "bold")
        ).pack(side="left", padx=10, pady=3)
        
        self.scroll_frame = ScrollableFrame(self.content_container, bg="#222222")
        self.scroll_frame.pack(expand=True, fill="both")
        
        # Reminder placeholder (row 1, 50% height)
        reminder_placeholder = tk.Frame(grid_container, bg="#111111")
        reminder_placeholder.grid(row=1, column=0, sticky="nsew", padx=0, pady=(2, 0))
        
        # Reminder section header
        reminder_header = tk.Frame(reminder_placeholder, bg="#333333", height=20)
        reminder_header.pack(fill="x", side="top")
        reminder_header.pack_propagate(False)  # Maintain fixed height
        
        tk.Label(
            reminder_header, text="Flagged/Reminders", 
            bg="#333333", fg="#AAAAAA",
            font=(self.font_family, 9, "bold")
        ).pack(side="left", padx=10, pady=3)
        
        tk.Label(
            reminder_placeholder, text="[Future Reminder Pane - 50% space]",
            bg="#111111", fg="#666666", font=(self.font_family, 9, "italic")
        ).pack(expand=True)

        # Resize Grip (Overlay on the right edge)
        self.resize_grip = tk.Frame(self.main_frame, bg="#666666", cursor="sb_h_double_arrow", width=5)
        self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
        self.resize_grip.bind("<B1-Motion>", self.on_resize_drag)
        self.resize_grip.bind("<ButtonRelease-1>", self.on_resize_release)

        # Hot Strip Visual overlay (only visible when collapsed)
        # We use a Canvas now to draw the animation
        self.hot_strip_canvas = tk.Canvas(self.main_frame, bg="#007ACC", highlightthickness=0)
        
        # --- Events ---
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Motion>", self.on_motion) 

        # Initial Load
        self.refresh_emails()
        
        # Initial State
        self.apply_state()

    def quit_application(self):
        """Terminates the application."""
        self.destroy()
        sys.exit(0)



    def open_settings(self):
        """Toggle the settings panel."""
        self.toggle_settings_panel()

    def toggle_settings_panel(self):
        """Show or hide the settings panel alongside the email list."""
        if self.settings_panel_open:
            # Close the panel
            if self.settings_panel:
                self.settings_panel.pack_forget()
                self.settings_panel.destroy()
                self.settings_panel = None
            self.settings_panel_open = False
            
            # Unfreeze main_frame so it can expand/contract with window resize
            self.main_frame.pack_propagate(True)
            self.main_frame.config(width=0)  # Remove fixed width
            self.main_frame.pack(side="left", fill="both", expand=True)
            
            # Restore original width
            self.set_geometry(self.expanded_width)
        else:
            # Freeze main_frame at its current width before expanding
            current_width = self.main_frame.winfo_width()
            self.main_frame.config(width=current_width)
            self.main_frame.pack_propagate(False)
            self.main_frame.pack(side="left", fill="y", expand=False)
            
            # Open the panel alongside email list
            self.settings_panel = SettingsPanel(self.content_wrapper, self, self.refresh_emails)
            self.settings_panel.pack(side="left", fill="y")
            self.settings_panel_open = True
            
            # Expand window by exactly +350px
            new_width = self.expanded_width + self.settings_panel_width
            self.set_geometry(new_width)
        
    def load_icon_white(self, path, size=None):
        """Loads an image, converts it to white, and returns ImageTk.PhotoImage."""
        try:
            pil_img = Image.open(path).convert("RGBA")
            
            # Resize if needed (optional, but good for consistency)
            if size:
                pil_img = pil_img.resize(size, Image.Resampling.LANCZOS)
                
            # Create a white image of the same size
            white_img = Image.new("RGBA", pil_img.size, (255, 255, 255, 255))
            
            # Boost Alpha: Treat any non-transparent pixel as fully opaque (or at least boost it)
            # This fixes "dim" icons that have low opacity
            # Boost Alpha based on user setting
            # If brightness is 1.0, threshold is 20. If 2.0 (max), threshold is lower (more sensitive) or we boost alpha values.
            # User wants "Control Brightness".
            # Simple approach: Multiply alpha by brightness factor.
            
            r, g, b, a = pil_img.split()
            
            r, g, b, a = pil_img.split()
            
            # Use static 75% brightness (grey)
            grey_val = 191
            white_img = Image.new("RGBA", pil_img.size, (grey_val, grey_val, grey_val, 255))
            
            # Simple threshold mask
            mask = a.point(lambda p: 255 if p > 20 else 0)
             
            # Use boosted mask
            final_img = Image.new("RGBA", pil_img.size, (0, 0, 0, 0))
            final_img.paste(white_img, (0, 0), mask=mask)
            
            return ImageTk.PhotoImage(final_img)
        except Exception as e:
            print(f"Error loading/coloring icon {path}: {e}")
            return None

    def handle_custom_action(self, config, email_data):
        """Executes the selected actions on the specific email."""
        print(f"Executing Actions for {config.get('label')} on {email_data.get('subject')}")
        
        entry_id = email_data.get("entry_id")
        if not entry_id:
            print("No EntryID found.")
            return

        item = self.outlook_client.get_item_by_entryid(entry_id)
        if not item:
            print("Could not retrieve Outlook item.")
            return

        # Sequential Execution Helper
        def execute_single_action(act_name, folder_name=""):
            if not act_name or act_name == "None": return
            
            try:
                if act_name == "Mark Read":
                    item.UnRead = False
                    item.Save()
                elif act_name == "Delete":
                    item.Delete()
                elif act_name == "Read & Delete":
                    item.UnRead = False
                    item.Save()
                    item.Delete()
                elif act_name == "Flag":
                    if item.IsMarkedAsTask: item.ClearTaskFlag()
                    else: item.MarkAsTask(4)
                    item.Save()
                elif act_name == "Open Email":
                    item.Display()
                    try:
                        # Maximize Window
                        inspector = item.GetInspector
                        inspector.WindowState = 2 # olMaximized
                        # Force window to front
                        inspector.Activate()
                    except:
                        pass
                elif act_name == "Reply":
                    # Mark as read first
                    item.UnRead = False
                    item.Save()
                    
                    reply = item.Reply()
                    reply.Display()
                    try:
                        # Maximize Window
                        inspector = reply.GetInspector
                        inspector.WindowState = 2 # olMaximized
                        inspector.Activate()
                    except:
                        pass
                elif act_name == "Move To...":
                    if folder_name:
                        target = self.outlook_client.find_folder_by_name(folder_name)
                        if target: item.Move(target)
                        else: print(f"Folder '{folder_name}' not found.")
            except Exception as e:
                print(f"Error executing {act_name}: {e}")

        try:
            # Execute Action 1
            execute_single_action(config.get("action1"), config.get("folder"))
            
            # Execute Action 2 - REMOVED
            # execute_single_action(config.get("action2"), config.get("folder"))
                
            # Refresh UI
            self.after(500, self.refresh_emails)
            
        except Exception as e:
            print(f"Action execution loop error: {e}")

    def toggle_card_actions(self, action_frame):
        if action_frame.winfo_viewable():
            action_frame.pack_forget()
        else:
            action_frame.pack(fill="x", pady=(5, 0))

    def start_polling(self):
        self.start_polling()
        
    def start_polling(self):
        """Poll Outlook every 30 seconds for new mail."""
        if self.outlook_client.check_new_mail():
            self.start_pulse()
            self.refresh_emails() # Auto-refresh list
            
        self.after(self.poll_interval * 1000, self.start_polling) # Dynamic interval
        
    def start_pulse(self):
        if not self.pulsing:
            self.pulsing = True
            self.pulse_step = 0
            self.run_pulse_animation()
            
    def stop_pulse(self):
        if self.pulsing:
            self.pulsing = False
            if self._pulse_job:
                self.after_cancel(self._pulse_job)
                self._pulse_job = None
            # Reset
            self.hot_strip_canvas.delete("pulse")

    def run_pulse_animation(self):
        if not self.pulsing: return
        
        # Calculate Height factor using sine wave (0.0 to 1.0)
        # math.sin goes from -1 to 1. We want 0 to 1 back to 0.
        # shifting phase to start at 0
        factor = (math.sin(self.pulse_step) + 1) / 2 # 0 to 1
        
        # Alternatively, for a "growth" from center:
        # We can just cycle 0 -> PI
        
        self.hot_strip_canvas.delete("pulse")
        
        w = self.hot_strip_width
        h = self.screen_height
        
        # Dynamic height based on factor
        # Let's make it grow to full height then shrink
        bar_height = h * factor
        
        # Center coords
        y1 = (h / 2) - (bar_height / 2)
        y2 = (h / 2) + (bar_height / 2)
        
        # Draw the "light" bar
        self.hot_strip_canvas.create_rectangle(
            0, y1, w, y2,
            fill=self.pulse_color,
            outline="",
            tags="pulse"
        )
        
        self.pulse_step += self.animation_speed
        
        # Speed: 50ms (20fps) for smooth gentle pulse
        self._pulse_job = self.after(50, self.run_pulse_animation)

    # --- Outlook Window Management (COM-based) ---

    def _get_outlook_app(self):
        """
        Gets the Outlook Application COM object.
        Tries GetActiveObject first (reuses existing instance), falls back to Dispatch.
        """
        try:
            # Try to connect to already-running Outlook
            app = win32com.client.GetActiveObject("Outlook.Application")
            return app
        except:
            pass
        
        # Fall back to Dispatch (may start Outlook if not running)
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            return app
        except Exception:
            return None

    def _get_any_explorer(self, app):
        """
        Returns an existing Explorer window if one exists, otherwise None.
        Tries ActiveExplorer first, then iterates Explorers collection.
        """
        if not app:
            return None
        
        # Try ActiveExplorer first
        try:
            explorer = app.ActiveExplorer()
            if explorer:
                return explorer
        except:
            pass
        
        # Iterate Explorers collection
        try:
            explorers = app.Explorers
            if explorers.Count > 0:
                return explorers.Item(1)
        except Exception:
            pass
        
        return None

    def _focus_window_by_hwnd(self, hwnd):
        """
        Brings a window to the foreground by its hwnd.
        Handles minimized windows and SetForegroundWindow restrictions.
        """
        if not hwnd:
            return False
        
        try:
            # Check if minimized
            if user32.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            
            # SetForegroundWindow can fail if our process doesn't have focus
            # Workaround: briefly set TOPMOST then remove it
            try:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except:
                pass
            
            win32gui.SetForegroundWindow(hwnd)
            return True
        except Exception:
            return False

    def _show_outlook_folder(self, folder_id):
        """
        Shows the specified Outlook folder (6=Inbox, 9=Calendar).
        Reuses existing Explorer if available, otherwise creates one.
        """
        try:
            app = self._get_outlook_app()
            if not app:
                return
            
            ns = app.GetNamespace("MAPI")
            folder = ns.GetDefaultFolder(folder_id)
            
            # Try to get existing explorer
            explorer = self._get_any_explorer(app)
            
            if explorer:
                # Reuse existing explorer - switch folder
                try:
                    explorer.CurrentFolder = folder
                except Exception:
                    pass
                
                # Get hwnd and focus
                try:
                    hwnd = explorer.Hwnd if hasattr(explorer, 'Hwnd') else None
                    if hwnd:
                        self._focus_window_by_hwnd(hwnd)
                    else:
                        explorer.Activate()
                except Exception:
                    pass
            else:
                # No explorer exists - create one via GetExplorer
                try:
                    new_explorer = folder.GetExplorer()
                    new_explorer.Display()
                    
                    # Focus the new window
                    self.after(100, lambda: self._focus_window_by_hwnd(
                        new_explorer.Hwnd if hasattr(new_explorer, 'Hwnd') else None
                    ))
                except Exception:
                    # Ultimate fallback
                    folder.Display()
                    
        except Exception:
            pass

    def open_outlook_app(self):
        """Opens/Focuses the main Outlook window (Inbox)."""
        self._show_outlook_folder(6)  # 6 = olFolderInbox

    def open_calendar_app(self):
        """Opens/Focuses the Outlook Calendar."""
        self._show_outlook_folder(9)  # 9 = olFolderCalendar
        
    def load_config(self):
        try:
            with open("sidebar_config.json", "r") as f:
                data = json.load(f)
                self.expanded_width = data.get("width", 300)
                self.is_pinned = data.get("pinned", True)
                self.dock_side = data.get("dock_side", "Left")
                self.font_family = data.get("font_family", "Segoe UI")
                self.font_size = data.get("font_size", 9)
                self.poll_interval = data.get("poll_interval", 30)
                self.btn_count = data.get("btn_count", 2)
                self.btn_config = data.get("btn_config", [
                    {"label": "Trash", "icon": "✕", "action1": "Mark Read", "action2": "Delete", "folder": ""}, 
                    {"label": "Reply", "icon": "↩", "action1": "Reply", "action2": "None", "folder": ""}
                ])
                self.show_read = data.get("show_read", False)
                self.only_flagged = data.get("only_flagged", False)
                self.include_read_flagged = data.get("include_read_flagged", True)
                self.flag_date_filter = data.get("flag_date_filter", "Anytime")
        except FileNotFoundError:
            pass

    def save_config(self):
        data = {
            "width": self.expanded_width,
            "pinned": self.is_pinned,
            "dock_side": self.dock_side,
            "font_family": self.font_family,
            "font_size": self.font_size,
            "poll_interval": self.poll_interval,
            "btn_count": self.btn_count,
            "btn_config": self.btn_config,
            "show_read": self.show_read,
            "only_flagged": self.only_flagged,
            "include_read_flagged": self.include_read_flagged,
            "flag_date_filter": self.flag_date_filter
        }
        with open("sidebar_config.json", "w") as f:
            json.dump(data, f)

    def refresh_emails(self):
        # Update UI fonts for header elements
        self.lbl_title.config(font=(self.font_family, 10, "bold"))
        self.btn_settings.config(font=(self.font_family, 12))
        self.btn_refresh.config(font=(self.font_family, 15))

        # Clear existing
        for widget in self.scroll_frame.scrollable_frame.winfo_children():
            widget.destroy()

        emails = self.outlook_client.get_inbox_items(
            count=30, 
            unread_only=not self.show_read
        )
        
        for email in emails:
            # Determine styling based on UnRead status
            is_unread = email.get('unread', False)
            bg_color = "#2d2d2d"
            # Blue border for unread, grey for read
            border_color = "#007ACC" if is_unread else "#555555"
            border_width = 2 if is_unread else 1
            
            # Create Card
            card = tk.Frame(
                self.scroll_frame.scrollable_frame, 
                bg=bg_color, 
                highlightbackground=border_color, 
                highlightthickness=border_width,
                padx=5, pady=5
            )
            card.pack(fill="x", expand=True, padx=2, pady=2)
            
            # --- Badge System (Follow-up Indicators) ---
            badge_text = ""
            badge_bg = "#555555" # Default
            
            if email.get('flag_status', 0) != 0:
                due = email.get('due_date')
                now_dt = datetime.now()
                received = email.get('received')
                
                # Check for 4501 "No Date"
                is_real_due = False
                if due:
                    try:
                        # Extract date part for comparison
                        due_short = due.replace(hour=0, minute=0, second=0, microsecond=0)
                        now_short = now_dt.replace(hour=0, minute=0, second=0, microsecond=0)
                        
                        if due_short.year < 3000: # Not the 4501 placeholder
                            is_real_due = True
                            diff = (due_short - now_short).days
                            
                            if diff < 0:
                                badge_text = "OVERDUE"
                                badge_bg = "#D83B01" # Dark Red/Orange
                            elif diff == 0:
                                badge_text = "DUE TODAY"
                                badge_bg = "#FF8C00" # Orange
                            elif diff == 1:
                                badge_text = "TOMORROW"
                                badge_bg = "#0078D4" # Blue
                            elif diff < 7:
                                badge_text = due_short.strftime("%a").upper()
                                badge_bg = "#00B7C3" # Teal
                            else:
                                badge_text = due_short.strftime("%d %b").upper()
                                badge_bg = "#666666"
                    except:
                        pass

                if not is_real_due and received:
                    # Show "Flagged X days ago"
                    try:
                        diff = (now_dt.astimezone() - received.astimezone()).days
                        if diff == 0:
                            badge_text = "FLAGGED TODAY"
                        else:
                            badge_text = f"FLAGGED {diff}D"
                        badge_bg = "#8E8E8E"
                    except:
                        pass

            header_frame = tk.Frame(card, bg=bg_color)
            header_frame.pack(fill="x")

            # Sender
            sender_text = email['sender']
            if is_unread:
                sender_text = "● " + sender_text # Add indicator dot
                
            lbl_sender = tk.Label(
                header_frame, 
                text=sender_text, 
                fg="white", 
                bg=bg_color, 
                font=(self.font_family, self.font_size, "bold"),
                anchor="w"
            )
            lbl_sender.pack(side="left", fill="x", expand=True)

            if badge_text:
                lbl_badge = tk.Label(
                    header_frame, 
                    text=badge_text, 
                    fg="white", 
                    bg=badge_bg, 
                    font=(self.font_family, self.font_size - 2, "bold"),
                    padx=6, pady=2
                )
                lbl_badge.pack(side="right", padx=2)
                
            # Subject
            lbl_subject = tk.Label(
                card, 
                text=email['subject'], 
                fg="#cccccc", 
                bg=bg_color, 
                font=(self.font_family, self.font_size),
                anchor="w",
                justify="left",
                wraplength=self.expanded_width - 40 
            )
            lbl_subject.pack(fill="x")
            
            # Preview (Body)
            lbl_preview = tk.Label(
                card, 
                text=email['preview'], 
                fg="#999999", 
                bg=bg_color, 
                font=(self.font_family, self.font_size - 1),
                anchor="w",
                justify="left",
                wraplength=self.expanded_width - 40 
            )
            lbl_preview.pack(fill="x")
            
            # --- Action Frame (Always Visible) ---
            action_frame = tk.Frame(card, bg=bg_color)
            action_frame.pack(fill="x", expand=True, padx=2, pady=(0, 2))
            
            # Add buttons to action_frame
            # Filter for valid buttons (Must have Icon AND Action)
            valid_buttons = [
                conf for conf in self.btn_config 
                if conf.get("icon") and conf.get("action1") != "None"
            ]
            
            # Limit to configured count if needed, but usually filtering is enough
            # valid_buttons = valid_buttons[:self.btn_count] 

            for conf in valid_buttons:
                icon = conf.get("icon", "🔘")
                
                is_png = icon.lower().endswith(".png")
                btn_image = None
                
                if is_png:
                    # Try to load from cache or disk
                    if icon in self.image_cache:
                        btn_image = self.image_cache[icon]
                    else:
                        path = os.path.join("icons", icon)
                        if os.path.exists(path):
                            # Load and color white, resize to ~24x24 for buttons (slightly bigger)
                            btn_image = self.load_icon_white(path, size=(24, 24))
                            if btn_image:
                                self.image_cache[icon] = btn_image
                
                if btn_image:
                    btn = tk.Label(
                        action_frame, 
                        image=btn_image, 
                        bg=bg_color,
                        padx=10, pady=5,
                        cursor="hand2"
                    )
                else:
                    btn = tk.Label(
                        action_frame, 
                        text=icon, 
                        fg="white", 
                        bg=bg_color,
                        font=("Segoe UI", 12),
                        padx=10, pady=5,
                        cursor="hand2"
                    )
                
                # If only 1 button, restrict width slightly to avoid massive button
                if len(valid_buttons) == 1:
                    btn.pack(side="left", expand=True, fill="y", ipadx=20)
                else:
                    btn.pack(side="left", expand=True, fill="both")
                
                # Bind hover
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#444444"))
                btn.bind("<Leave>", lambda e, b=btn, bg=bg_color: b.config(bg=bg))
                
                # Tooltip logic
                act1 = conf.get("action1", "")
                act2 = conf.get("action2", "None")
                if act2 != "None":
                    tip_text = f"{act1} & {act2}"
                else:
                    tip_text = act1
                
                ToolTip(btn, tip_text)
                
                # Click handler
                btn.bind("<Button-1>", lambda e, c=conf, em=email: self.handle_custom_action(c, em))
            
            # --- Bindings Removed ---
            # card.bind("<Button-1>", ...) 
            
            # Dynamic wrapping for both labels
            def update_wraps(e, s=lbl_subject, p=lbl_preview):
                width = e.width - 20
                s.config(wraplength=width)
                p.config(wraplength=width)
                
            card.bind("<Configure>", update_wraps)

    def draw_pin_icon(self):
        self.btn_pin.delete("all")
        color = "#007ACC" if self.is_pinned else "#AAAAAA"
        # Draw a simple pin shape
        self.btn_pin.create_oval(10, 5, 20, 15, fill=color, outline="")
        self.btn_pin.create_line(15, 15, 15, 25, fill=color, width=2)

    def toggle_pin(self):
        self.is_pinned = not self.is_pinned
        self.draw_pin_icon()
        self.save_config()
        self.apply_state()

    def apply_state(self):
        """Applies the current state (Pinned/Expanded/Collapsed) to the window and AppBar."""
        # Update monitor info to ensure correct sizing on monitor change
        self.monitor_x, self.monitor_y, self.screen_width, self.screen_height = self.get_current_monitor_info()

        # Update AppBar edge based on preference
        new_edge = ABE_LEFT if self.dock_side == "Left" else ABE_RIGHT
        
        # If side changed, we MUST unregister the old one first to release the old edge
        if self.appbar.edge != new_edge:
            self.appbar.unregister()
            self.appbar.edge = new_edge
            self.appbar.abd.uEdge = new_edge

        if self.is_pinned:
            # Pinned: Always Expanded, Always Reserved (Docked)
            self.hot_strip_canvas.place_forget()
            self.header.pack(fill="x", side="top")
            # self.content_container.pack(expand=True, fill="both", padx=5, pady=5)  # Now managed by grid
            
            # Place grip on opposite side of dock
            if self.dock_side == "Left":
                self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
            else:
                self.resize_grip.place(relx=0.0, rely=0, anchor="nw", relheight=1.0)
            
            self.appbar.register() # This will re-register on the new edge
            # Use authoritative position from Windows to avoid gaps
            x, y, w, h = self.appbar.set_pos(self.expanded_width, self.monitor_x, self.monitor_y, self.screen_width, self.screen_height)
            self.geometry(f"{w}x{h}+{x}+{y}")
            self.update_idletasks()
            self.is_expanded = True
            
        elif self.is_expanded:
            # Expanded (Hover): Broad width, BUT acts as OVERLAY (No docking/reservation)
            self.hot_strip_canvas.place_forget()
            self.header.pack(fill="x", side="top")
            # For overlay mode, we still show the content
            # self.content_container.pack(expand=True, fill="both", padx=5, pady=5)  # Now managed by grid
            
            if self.dock_side == "Left":
                self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
            else:
                self.resize_grip.place(relx=0.0, rely=0, anchor="nw", relheight=1.0)
            
            # Unregister AppBar so we don't push other windows
            self.appbar.unregister()
            
            self.set_geometry(self.expanded_width)
            
        else:
            # Collapsed: Thin width, Overlay
            self.appbar.unregister() # Release space
            
            # Hide internals to prevent squishing
            self.header.pack_forget()
            # self.content_container.pack_forget()  # Now managed by grid
            self.resize_grip.place_forget()
            
            # Show Hot Strip
            self.hot_strip_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
            
            self.set_geometry(self.hot_strip_width)

    def on_resize_drag(self, event):
        if self.is_pinned or self.is_expanded:
            x_root = self.winfo_pointerx()
            
            # Calculate width based on side
            if self.dock_side == "Left":
                new_width = x_root - self.monitor_x
            else:
                new_width = (self.monitor_x + self.screen_width) - x_root
            
            if new_width > self.min_width and new_width < (self.screen_width // 2):
                self.expanded_width = new_width
                # Optimization: ONLY resize the visual window, do NOT trigger AppBar reflow
                self.set_geometry(self.expanded_width)
                # Ensure the content knows we resized if needed (pack handles this)

    def on_resize_release(self, event):
        # Commit the new width to the system (triggers reflow once)
        self.apply_state() 
        self.save_config()

    def set_geometry(self, width):
        # Always dock to preferred side, full height of CURRENT screen
        mx, my, mw, mh = self.get_current_monitor_info()
        
        if self.dock_side == "Left":
            x = mx
        else:
            x = mx + mw - width
            
        self.geometry(f"{width}x{mh}+{x}+{my}")
        # Ensure window updates its position immediately
        self.update_idletasks()
        # Force top most again just in case
        self.wm_attributes("-topmost", True)

    def get_current_monitor_info(self):
        """Retrieves the geometry of the monitor closest to the window center."""
        hwnd = self.winfo_id()
        # Ensure we have the actual top-level window handle
        hwnd = ctypes.windll.user32.GetParent(hwnd) or hwnd
        monitor = user32.MonitorFromWindow(hwnd, 2) # MONITOR_DEFAULTTONEAREST
        
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if user32.GetMonitorInfoW(monitor, ctypes.byref(mi)):
            return (mi.rcMonitor.left, mi.rcMonitor.top, 
                    mi.rcMonitor.right - mi.rcMonitor.left, 
                    mi.rcMonitor.bottom - mi.rcMonitor.top)
            
        # Fallback to defaults
        return (0, 0, self.winfo_screenwidth(), self.winfo_screenheight())

    def start_window_drag(self, event):
        self._win_drag_x = event.x
        self._win_drag_y = event.y
        # Temporarily unregister AppBar so we can move freely
        self.appbar.unregister()

    def on_window_drag(self, event):
        deltax = event.x - self._win_drag_x
        deltay = event.y - self._win_drag_y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        # During drag, we don't snap/resize, just move
        self.geometry(f"+{x}+{y}")

    def stop_window_drag(self, event):
        # Auto-Snap Logic
        mx, my, mw, mh = self.get_current_monitor_info()
        
        # Calculate window center
        win_x = self.winfo_x()
        win_w = self.winfo_width()
        win_center = win_x + (win_w / 2)
        
        # Monitor center
        mon_center = mx + (mw / 2)
        
        # Determine side
        if win_center < mon_center:
            self.dock_side = "Left"
        else:
            self.dock_side = "Right"
            
        # Re-apply state which will snap to monitor edge and re-register
        self.apply_state()


    def on_enter(self, event):
        # Stop pulsing on interaction
        self.stop_pulse()
        
        if self._collapse_timer:
            self.after_cancel(self._collapse_timer)
            self._collapse_timer = None
        
        if not self.is_pinned and not self.is_expanded:
            self.is_expanded = True
            self.apply_state() # Expand and reserve space

    def on_leave(self, event):
        # We need to be careful. Leaving the window to the desktop should collapse.
        # But verify we aren't just hovering a child widget (Tkinter events bubble, but checking coordinates keeps us safe).
        x, y = self.winfo_pointerxy()
        widget_under_mouse = self.winfo_containing(x, y)
        
        # If we are really outside the window
        if not self.is_pinned and self.is_expanded:
             # Delay collapse
             if self._collapse_timer:
                 self.after_cancel(self._collapse_timer)
             self._collapse_timer = self.after(self.hover_delay, self.do_collapse)

    def on_motion(self, event):
        # Reset collapse timer if moving inside
        if self._collapse_timer:
             self.after_cancel(self._collapse_timer)
             self._collapse_timer = None

    def do_collapse(self):
        if not self.is_pinned:
            self.is_expanded = False
            self.apply_state() # Collapse and release space

    def get_app_data_dir(self):
        """Returns the appropriate application data directory."""
        # Use %LOCALAPPDATA% (Preferred for modern Windows apps)
        local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        app_dir = os.path.join(local_app_data, "OutlookSidebar")
        
        if not os.path.exists(app_dir):
            try:
                os.makedirs(app_dir)
            except OSError as e:
                print(f"Error creating app data dir: {e}")
                # Fallback to temp if strictly needed, or just fail
        return app_dir

    def load_config(self):
        try:
            app_dir = self.get_app_data_dir()
            config_path = os.path.join(app_dir, "config.json")
            
            # If user config doesn't exist, check for bundled default in current dir (read-only)
            if not os.path.exists(config_path):
                bundled_config = "config.json"
                if os.path.exists(bundled_config):
                    try:
                        # Copy bundled default to AppData so user can edit it later
                        shutil.copy2(bundled_config, config_path)
                    except Exception as e:
                        print(f"Failed to copy default config: {e}")
            
            if os.path.exists(config_path):
                with open(config_path, "r") as f:
                    config = json.load(f)
                    
                self.dock_side = config.get("dock_side", "Right")
                self.font_family = config.get("font_family", "Segoe UI")
                self.font_size = config.get("font_size", 9)
                self.poll_interval = config.get("poll_interval", 30)
                
                if "buttons" in config:
                     self.btn_config = config["buttons"]
                     self.btn_count = len(self.btn_config)
        except Exception as e:
            print(f"Error loading config: {e}")

    def save_config(self):
        app_dir = self.get_app_data_dir()
        config_path = os.path.join(app_dir, "config.json")
        
        config = {
            "dock_side": self.dock_side,
            "font_family": self.font_family,
            "font_size": self.font_size,
            "poll_interval": self.poll_interval,
            "buttons": self.btn_config
        }
        try:
            with open(config_path, "w") as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

# --- Single Instance Logic (Mutex) ---

class SingleInstance:
    """
    Limits application to a single instance using a Named Mutex.
    Safe for MSIX and standard execution.
    """
    def __init__(self, name="Global\\OutlookSidebar_Mutex_v1"):
        self.mutex_name = name
        self.mutex_handle = None
        self.last_error = 0

    def already_running(self):
        # CreateMutexW will return a handle. If it already existed, GetLastError returns ERROR_ALREADY_EXISTS
        ERROR_ALREADY_EXISTS = 183
        
        self.mutex_handle = kernel32.CreateMutexW(None, False, self.mutex_name)
        self.last_error = kernel32.GetLastError()
        
        if self.last_error == ERROR_ALREADY_EXISTS:
            return True
        return False
        
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self.mutex_handle:
            kernel32.CloseHandle(self.mutex_handle)

if __name__ == "__main__":
    # Check Single Instance
    app_instance = SingleInstance()
    if app_instance.already_running():
        # Optional: Bring existing window to front (Requires FindWindow/SetForegroundWindow logic)
        # For now, just exit silently or print
        # messagebox.showinfo("Outlook Sidebar", "The application is already running.")
        sys.exit(0)

    # Keep the mutex handle alive for the duration of the app
    app = SidebarWindow()
    app.mainloop()

