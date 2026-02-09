# -*- coding: utf-8 -*-
try:
    # Python 2
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
except ImportError:
    # Python 3
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

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
# from tkinter import messagebox (handled above)
from PIL import Image, ImageTk, ImageDraw
try:
    # Pillow 10+
    RESAMPLE_MODE = Image.Resampling.LANCZOS
except AttributeError:
    # Older Pillow
    RESAMPLE_MODE = Image.ANTIALIAS

from datetime import datetime, timedelta

# --- Store Compatibility Imports ---
import sys
import shutil
# Using ctypes for Mutex to avoid extra pywin32 module dependencies if not strictly needed,
# though win32event is also fine since win32gui is used.
# sticking to ctypes kernel32 for zero-dependency bloat for this specific feature.
kernel32 = ctypes.windll.kernel32


# --- Application Constants ---
VERSION = "v1.3.3"


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
        
        # 2. Adjust if necessary (System might have changed it)
        if self.edge == ABE_LEFT:
            self.abd.rc.right = self.abd.rc.left + width
        elif self.edge == ABE_RIGHT:
            self.abd.rc.left = self.abd.rc.right - width
        
        # 3. Set Position
        shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(self.abd))
        
        # DEBUG: Print committed rect
        # print(f"AppBar SetPos: Edge={self.edge}, Rect=({self.abd.rc.left}, {self.abd.rc.top}, {self.abd.rc.right}, {self.abd.rc.bottom})")

        # Return the actual rectangle committed
        return self.abd.rc.left, self.abd.rc.top, self.abd.rc.right - self.abd.rc.left, self.abd.rc.bottom - self.abd.rc.top

class ScrollableFrame(tk.Frame):
    """
    A scrollable frame that can contain multiple email cards.
    """
    def __init__(self, container, *args, **kwargs):
        tk.Frame.__init__(self, container, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=kwargs.get("bg", "#222222"), highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas, bg=kwargs.get("bg", "#222222"))

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        # Speed up scrolling (Windows default is slow)
        self.canvas.configure(yscrollincrement=5)

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self._scroll_timer = None

        # Custom Scroll Buttons
        # We place them relative to 'self' so they overlay the canvas
        self.btn_up = tk.Button(self, text="▲",  
                                bg="#444444", fg="white", bd=0, font=("Arial", 6), width=10, activebackground="#555555", activeforeground="white")
        self.btn_up.bind("<ButtonPress-1>", lambda e: self.start_scroll(-1))
        self.btn_up.bind("<ButtonRelease-1>", self.stop_scroll)
        self.btn_up.bind("<Leave>", self.stop_scroll)

        self.btn_down = tk.Button(self, text="▼", 
                                  bg="#444444", fg="white", bd=0, font=("Arial", 6), width=10, activebackground="#555555", activeforeground="white")
        self.btn_down.bind("<ButtonPress-1>", lambda e: self.start_scroll(1))
        self.btn_down.bind("<ButtonRelease-1>", self.stop_scroll)
        self.btn_down.bind("<Leave>", self.stop_scroll)

        self.canvas.configure(yscrollcommand=self._on_scroll_update)
        self.canvas.pack(side="left", fill="both", expand=True)
        # self.scrollbar removed
        
        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Ensure scrollable frame matches canvas width
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def start_scroll(self, direction):
        self.scroll(direction)
        # Initial delay before continuous scroll (e.g. 500ms)
        self._scroll_timer = self.after(300, lambda: self.auto_scroll(direction))

    def auto_scroll(self, direction):
        self.scroll(direction)
        # Repeat interval (e.g. 50ms)
        self._scroll_timer = self.after(50, lambda: self.auto_scroll(direction))

    def stop_scroll(self, event=None):
        if self._scroll_timer:
            self.after_cancel(self._scroll_timer)
            self._scroll_timer = None

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
            
        tw.wm_geometry("+{}+{}".format(x, y))

    def hide_tip(self):
        """Hides the tooltip."""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class OutlookClient:
    # Color Enum Logic (1-25)
    # Approximate Hex values for Dark Mode
    OL_CAT_COLORS = {
        0: "#555555", # None
        1: "#DA5758", # Red
        2: "#E68D49", # Orange
        3: "#EAC389", # Peach
        4: "#F0E16C", # Yellow
        5: "#81C672", # Green
        6: "#61CED1", # Teal
        7: "#97CD9B", # Olive
        8: "#6E93E6", # Blue
        9: "#A580DA", # Purple
        10: "#CE7091", # Maroon
        11: "#8BB2C2", # Steel
        12: "#6A8591", # Dark Steel
        13: "#ACACAC", # Gray
        14: "#6E6E6E", # Dark Gray
        15: "#333333", # Black
        16: "#BE4250", # Dark Red
        17: "#CA7532", # Dark Orange
        18: "#BD934A", # Dark Peach
        19: "#BDB84B", # Dark Yellow
        20: "#5E9348", # Dark Green
        21: "#3E9EA2", # Dark Teal
        22: "#689E59", # Dark Olive
        23: "#4566B0", # Dark Blue
        24: "#7750A8", # Dark Purple
        25: "#A14868",  # Dark Maroon
    }

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
            print("Error connecting to Outlook: {}".format(e))
            self.outlook = None
            self.namespace = None
            return False

    def get_accounts(self):
        """Returns list of account names."""
        accounts = []
        if not self.connect(): return []
        try:
            for store in self.namespace.Stores:
                accounts.append(store.DisplayName)
        except Exception as e:
            print("Error fetching accounts: {}".format(e))
        return accounts

    def _get_enabled_stores(self, account_names):
        """Helper: Yields stores that match the provided names (or all if None)."""
        if not self.namespace: return
        
        # If no specific accounts provided, try to assume all or default
        # But for strictly obeying 'enabled_accounts', we expect a list.
        # If None, we default to ALL stores (legacy behavior compatible)
        
        try:
            for store in self.namespace.Stores:
                if account_names is None or store.DisplayName in account_names:
                    yield store
        except Exception:
            return

    def check_latest_time(self, account_names=None):
        """Updates the globally tracked last_received_time from enabled accounts using safe Tables."""
        if not self.namespace: return
        
        latest = None
        
        try:
            for store in self._get_enabled_stores(account_names):
                try:
                    inbox = store.GetDefaultFolder(6)
                    # Use Table to avoid traversing MailItem objects (Security Guard)
                    table = inbox.GetTable()
                    table.Sort("ReceivedTime", True) # Descending
                    table.Columns.RemoveAll()
                    table.Columns.Add("ReceivedTime")
                    
                    if not table.EndOfTable:
                        row = table.GetNextRow()
                        if row:
                            # Use GetValues for safety
                            vals = row.GetValues()
                            t = vals[0]
                            if latest is None or t > latest:
                                latest = t
                except:
                    continue
                    
            if latest:
                self.last_received_time = latest
                
        except Exception:
             pass

    def check_new_mail(self, account_names=None):
        """Checks for new mail across enabled accounts using safe Tables."""
        for attempt in range(2):
            if not self.namespace:
                if not self.connect(): return False

            try:
                found_new = False
                global_max = self.last_received_time
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        inbox = store.GetDefaultFolder(6)
                        # Use Table for security safety
                        table = inbox.GetTable()
                        table.Sort("ReceivedTime", True)
                        table.Columns.RemoveAll()
                        table.Columns.Add("ReceivedTime")
                        
                        if not table.EndOfTable:
                            row = table.GetNextRow()
                            if row:
                                vals = row.GetValues()
                                current_time = vals[0]
                                # Compare against our global max
                                if self.last_received_time and current_time > self.last_received_time:
                                    # print(f"DEBUG: Found NEW mail! {current_time} > {self.last_received_time}")
                                    found_new = True
                                
                                # Update local tracker for this poll
                                if global_max is None or current_time > global_max:
                                    global_max = current_time
                    except Exception as e:
                        # print(f"DEBUG: Error checking store: {e}")
                        continue
                
                if global_max:
                    self.last_received_time = global_max
                    
                return found_new
                
            except Exception as e:
                print("Polling error (Attempt {}): {}".format(attempt+1, e))
                self.namespace = None 
        
        return False

    def get_unread_count(self, account_names=None, account_config=None):
        """Returns total unread count from configured folders."""
        total = 0
        if not self.namespace: 
            if not self.connect(): return 0
            
        for store in self._get_enabled_stores(account_names):
            try:
                folders_to_scan = []
                # Check config
                if account_config and store.DisplayName in account_config:
                    conf = account_config[store.DisplayName]
                    if "email_folders" in conf and conf["email_folders"]:
                        for path in conf["email_folders"]:
                            f = self.get_folder_by_path(store, path)
                            if f: folders_to_scan.append(f)
                
                # Fallback to Inbox
                if not folders_to_scan:
                    try: folders_to_scan.append(store.GetDefaultFolder(6))
                    except: pass
                    
                for f in folders_to_scan:
                    try: total += f.UnReadItemCount
                    except: pass
            except: continue
        return total

    def get_folder_by_path(self, store, folder_path):
        """Helper to navigate folder path string (e.g. 'Inbox/Subfolder')."""
        try:
            parts = folder_path.split("/")
            curr = store.GetRootFolder()
            for p in parts:
                curr = curr.Folders[p]
            return curr
        except:
            return None

    def get_calendar_items(self, start_dt, end_dt, account_names=None):
        """Fetches calendar items from all enabled accounts. Accepts datetime objects."""
        # Print Debug Log
        # Print Debug Log
        try:
             # print("DEBUG: Calendar Query: {} to {}".format(start_dt, end_dt))
             pass
        except: pass

        for attempt in range(2):
            if not self.namespace:
                 if not self.connect(): return []
            try:
                all_results = []
                
                # Format for DASL/Jet - US Format usually preferred
                s_str = start_dt.strftime('%m/%d/%Y %H:%M %p')
                e_str = end_dt.strftime('%m/%d/%Y %H:%M %p')
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        cal = store.GetDefaultFolder(9)
                        items = cal.Items
                        items.Sort("[Start]")
                        items.IncludeRecurrences = True
                        
                        restrict = "[Start] >= '{}' AND [Start] <= '{}'".format(s_str, e_str)
                        
                        try:
                            items = items.Restrict(restrict)
                        except Exception as e:
                            print("Restrict Warning: {}".format(e))
                            # Continue with full items (will be filtered manually below) NO, too slow.
                            # If restrict fails, maybe just skip or try simple filter?
                            pass
                        
                        for item in items:
                            try:
                                # Manual Date Check (Safety Net against locale issues)
                                # Normalize Item Start (Aware/Naive)
                                i_start = item.Start
                                if getattr(i_start, "tzinfo", None) is not None:
                                     i_start = i_start.replace(tzinfo=None) # Make naive for comparison with our naive start_dt/end_dt
                                
                                if i_start < start_dt or i_start > end_dt:
                                     continue

                                all_results.append({
                                    "subject": item.Subject,
                                    "start": item.Start,
                                    "location": getattr(item, "Location", ""),
                                    "entry_id": item.EntryID,
                                    "is_meeting": True,
                                    "response_status": getattr(item, "ResponseStatus", 0),
                                    "account": store.DisplayName # Optional: Track source
                                })
                            except:
                                continue
                    except:
                        continue
                        
                # Sort merged results by start time
                try:
                    all_results.sort(key=lambda x: x["start"])
                except:
                    pass
                    
                return all_results
            except Exception as e:
                print("Calendar error: {}".format(e))
                self.namespace = None
        return []

    def get_tasks(self, due_filters=None, account_names=None):
        """Fetches Outlook Tasks from enabled accounts using safe Tables."""
        for attempt in range(2):
            if not self.namespace:
                 if not self.connect(): return []
            try:
                all_results = []
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        tasks_folder = store.GetDefaultFolder(13)
                        
                        restricts = ["[Complete] = False"]
                        
                        # Date Filter Logic
                        if due_filters and len(due_filters) > 0:
                            date_queries = []
                            now = datetime.now()
                            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
                            tomorrow = today + timedelta(days=1)
                            db_tomorrow = today + timedelta(days=2)
                            
                            for filter_name in due_filters:
                                if filter_name == "Overdue":
                                    date_queries.append("[DueDate] < '{}'".format(today.strftime('%m/%d/%Y %I:%M %p'))) 
                                elif filter_name == "Today":
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(today.strftime('%m/%d/%Y %I:%M %p'), tomorrow.strftime('%m/%d/%Y %I:%M %p')))
                                elif filter_name == "Tomorrow":
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(tomorrow.strftime('%m/%d/%Y %I:%M %p'), db_tomorrow.strftime('%m/%d/%Y %I:%M %p')))
                                elif filter_name == "Next 7 Days":
                                    next_week = today + timedelta(days=8)
                                    date_queries.append("([DueDate] >= '{}' AND [DueDate] < '{}')".format(today.strftime('%m/%d/%Y %I:%M %p'), next_week.strftime('%m/%d/%Y %I:%M %p')))
                                elif filter_name == "No Date":
                                    date_queries.append("([DueDate] IS NULL OR [DueDate] > '01/01/4500')")
                            
                            if date_queries:
                                combined_date_query = " OR ".join(date_queries)
                                restricts.append("({})".format(combined_date_query))

                        restrict_str = " AND ".join(restricts) if restricts else ""
                        
                        try:
                            table = tasks_folder.GetTable(restrict_str) if restrict_str else tasks_folder.GetTable()
                        except:
                            continue

                        table.Columns.RemoveAll()
                        table.Columns.Add("Subject")
                        table.Columns.Add("DueDate")
                        table.Columns.Add("EntryID")
                        
                        count = 0
                        while not table.EndOfTable and count < 30:
                            row = table.GetNextRow()
                            if not row: break
                            
                            try:
                                vals = row.GetValues()
                                
                                all_results.append({
                                    "subject": vals[0],
                                    "due": vals[1],
                                    "entry_id": vals[2],
                                    "is_task": True,
                                    "account": store.DisplayName
                                })
                                count += 1
                            except:
                                continue
                    except:
                        continue # Skip store if tasks failed
                        
                # Sort combined results
                all_results.sort(key=lambda x: x["due"].timestamp() if getattr(x["due"], 'timestamp', None) else 0)
                
                return all_results
            except Exception as e:
                print("Tasks error: {}".format(e))
                self.namespace = None
        return []

    def get_inbox_items(self, count=20, unread_only=False, only_flagged=False, due_filters=None, account_names=None, account_config=None):
        """Fetches items from configured folders for enabled accounts."""
        for attempt in range(2):
            if not self.namespace:
                if not self.connect(): return [], 0

            try:
                all_items = []
                total_unread_count = 0
                
                for store in self._get_enabled_stores(account_names):
                    try:
                        # Determine folders to scan
                        folders_to_scan = []
                        
                        # Check config for this account
                        if account_config and store.DisplayName in account_config:
                            conf = account_config[store.DisplayName]
                            if "email_folders" in conf and conf["email_folders"]:
                                for path in conf["email_folders"]:
                                    f = self.get_folder_by_path(store, path)
                                    if f: folders_to_scan.append(f)
                        
                        # Fallback to Inbox if no specific folders configured
                        if not folders_to_scan:
                            try:
                                folders_to_scan.append(store.GetDefaultFolder(6))
                            except: pass
                            
                        for folder in folders_to_scan:
                             try:
                                 total_unread_count += folder.UnReadItemCount
                             except: pass
                             items = self._fetch_items_from_inbox_folder(folder, count, unread_only, only_flagged, due_filters, store)
                             all_items.extend(items)
                    except:
                        continue
                        
                # Sort merged results by ReceivedTime (Descending)
                # Sort merged results by ReceivedTime (Descending)
                def sort_key(x):
                    dt = x.get("received_dt")
                    if dt:
                        return dt
                    return datetime.min

                all_items.sort(key=sort_key, reverse=True)
                
                return all_items[:count], total_unread_count
                
            except Exception as e:
                print("Inbox error: {}".format(e))
                
        return [], 0

    def _fetch_items_from_inbox_folder(self, folder, count, unread_only, only_flagged, due_filters, store):
        """Helper to fetch items from a single inbox folder."""
        restricts = []
        
        if only_flagged:
            restricts.append("[FlagStatus] <> 0")
            
            if due_filters and len(due_filters) > 0:
                date_queries = []
                now = datetime.now()
                today = now.replace(hour=0, minute=0, second=0, microsecond=0)
                tomorrow = today + timedelta(days=1)
                db_tomorrow = today + timedelta(days=2)
                
                for filter_name in due_filters:
                    if filter_name == "Overdue":
                        date_queries.append("[TaskDueDate] < '{}'".format(today.strftime('%m/%d/%Y %I:%M %p'))) 
                    elif filter_name == "Today":
                        date_queries.append("([TaskDueDate] >= '{}' AND [TaskDueDate] < '{}')".format(today.strftime('%m/%d/%Y %I:%M %p'), tomorrow.strftime('%m/%d/%Y %I:%M %p')))
                    elif filter_name == "Tomorrow":
                        date_queries.append("([TaskDueDate] >= '{}' AND [TaskDueDate] < '{}')".format(tomorrow.strftime('%m/%d/%Y %I:%M %p'), db_tomorrow.strftime('%m/%d/%Y %I:%M %p')))
                    elif filter_name == "Next 7 Days":
                        next_week = today + timedelta(days=8)
                        date_queries.append("([TaskDueDate] >= '{}' AND [TaskDueDate] < '{}')".format(today.strftime('%m/%d/%Y %I:%M %p'), next_week.strftime('%m/%d/%Y %I:%M %p')))
                    elif filter_name == "No Date":
                        date_queries.append("([TaskDueDate] IS NULL OR [TaskDueDate] > '01/01/4500')")
                
                if date_queries:
                    combined_date_query = " OR ".join(date_queries)
                    restricts.append("({})".format(combined_date_query))
        else:
            if unread_only:
                restricts.append("[UnRead] = True")
        
        restrict_str = " AND ".join(restricts) if restricts else ""
        
        try:
            table = folder.GetTable(restrict_str) if restrict_str else folder.GetTable()
        except:
            return []

        table.Columns.RemoveAll()
        
        desired_cols = [
            "EntryID", "Subject", "SenderName", "ReceivedTime", 
            "UnRead", "FlagStatus", "TaskDueDate", "Importance", "Categories"
        ]
        
        active_cols = []
        for c in desired_cols:
            try: 
                table.Columns.Add(c)
                active_cols.append(c)
            except: pass
            
        # Optional Props
        attach_prop = "urn:schemas:httpmail:hasattachment"
        try: 
            table.Columns.Add(attach_prop)
            active_cols.append("has_attachments_prop")
        except: pass

        body_prop = "http://schemas.microsoft.com/mapi/proptag/0x1000001E"
        try: 
            table.Columns.Add(body_prop)
            active_cols.append("body_prop")
        except: pass
        
        try:
            table.Sort("ReceivedTime", True) 
        except: pass
        
        results = []
        while not table.EndOfTable and len(results) < count:
            row = table.GetNextRow()
            if not row: break
            
            try:
                vals = row.GetValues()
                
                item_data = {}
                if vals:
                     if len(vals) != len(active_cols):
                         # Attempt to map anyway up to min length
                         limit = min(len(vals), len(active_cols))
                         for i in range(limit):
                             item_data[active_cols[i]] = vals[i]
                     else:
                        for i, col_name in enumerate(active_cols):
                            item_data[col_name] = vals[i]
                
                # Normalize Data
                received_dt = item_data.get("ReceivedTime")
                received_str = ""
                if received_dt:
                    try:
                        # PyTime compatible conversion
                        if hasattr(received_dt, "strftime"):
                            received_str = received_dt.strftime("%d/%m %H:%M")
                        else:
                            # Fallback for PyTime objects (Python 2.7 win32com)
                            # Casting to str normally gives "YYYY-MM-DD HH:MM:SS..."
                            # Or convert to datetime via timestamp if available, or just explicit construction
                            try:
                                # Start can be accessed via .year, .month etc if it's PyTime
                                d = datetime(received_dt.year, received_dt.month, received_dt.day, received_dt.hour, received_dt.minute, received_dt.second)
                                received_str = d.strftime("%d/%m %H:%M")
                                # Replace the PyTime object with the python datetime object for later use
                                item_data["ReceivedTime"] = d
                            except:
                                received_str = str(received_dt)[:16] # Fallback
                                
                    except Exception as e:
                        print("DEBUG: Date conversion error: {}".format(e))
                        received_str = str(received_dt)
                    
                entry_id = item_data.get("EntryID")
                
                # Construct result object
                res = {
                    "subject": item_data.get("Subject", "(No Subject)"),
                    "sender": item_data.get("SenderName", "Unknown"),
                    "received": received_str,
                    "received_dt": received_dt,
                    "unread": item_data.get("UnRead", False),
                    "has_attachment": item_data.get("has_attachments_prop", False),
                    "flag_status": item_data.get("FlagStatus", 0),
                    "due_date": item_data.get("TaskDueDate"),
                    "importance": item_data.get("Importance", 1), # Default 1 (Normal)
                    "categories": item_data.get("Categories", ""),
                    "body": item_data.get("body_prop", "")[:3000] if item_data.get("body_prop") else "",
                    "entry_id": entry_id,
                    "store_id": store.StoreID, # New field
                    "account": store.DisplayName
                }
                results.append(res)
            except Exception as e:
                print("DEBUG: Row processing error: {}".format(e))
                continue
                
        return results


    def get_item_by_entryid(self, entry_id, store_id=None):
        """Retrieves a specific Outlook item by its EntryID."""
        if not self.namespace:
            self.connect()
        try:
            if store_id:
                return self.namespace.GetItemFromID(entry_id, store_id)
            return self.namespace.GetItemFromID(entry_id)
        except Exception as e:
            print("Error getting item {}: {}".format(entry_id, e))
            return None

    def get_folder_by_path(self, store, path_str):
        """Resolves a folder path (e.g., 'Inbox/ProjectX') to a MAPIFolder object."""
        if not path_str: return None
        
        try:
            parts = path_str.split("/")
            # Start at root of the store
            current = store.GetRootFolder()
            
            for part in parts:
                found = False
                for f in current.Folders:
                    if f.Name == part:
                        current = f
                        found = True
                        break
                if not found:
                    return None
            return current
        except Exception as e:
            print("Error resolving path '{}': {}".format(path_str, e))
            return None

    def mark_task_complete(self, entry_id, store_id=None):
        """Marks a task as complete."""
        try:
            item = self.get_item_by_entryid(entry_id, store_id)
            if item:
                item.MarkComplete()
                item.Save()
                return True
        except Exception as e:
            print("Error marking task complete: {}".format(e))
            return False

    def dismiss_calendar_item(self, entry_id):
        """Dismisses (Deletes) a calendar item."""
        try:
            item = self.get_item_by_entryid(entry_id)
            if item:
                item.Delete()
                return True
        except Exception as e:
            print("Error dismissing calendar item: {}".format(e))
            return False

    def unflag_email(self, entry_id, store_id=None):
        """Unflags an email (marks as complete)."""
        try:
            item = self.get_item_by_entryid(entry_id, store_id)
            if item:
                # 1 = olFlagComplete
                item.FlagStatus = 1 
                item.Save()
                return True
        except Exception as e:
            print("Error unflagging email: {}".format(e))
            return False
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
            print("Error finding folder {}: {}".format(folder_name, e))
            return None


            
    def get_folder_list(self, account_name=None):
        """Returns a list of folder paths. If account_name provided, scans that store."""
        if not self.namespace: 
            self.connect()
            
        folders = []
        try:
            root_folder = None
            if account_name:
                for store in self.namespace.Stores:
                    if store.DisplayName == account_name:
                        root_folder = store.GetRootFolder()
                        break
            else:
                 # Default logic (first account)
                 root_folder = self.namespace.GetDefaultFolder(6).Parent

            if not root_folder: return []
            
            def recurse(folder, parent_path=""):
                try:
                    name = folder.Name
                    path = "{}/{}".format(parent_path, name) if parent_path else name
                    folders.append(path)
                    
                    # 3 levels deep max
                    if parent_path.count("/") < 3:
                        for sub in folder.Folders:
                            recurse(sub, path)
                except: pass

            for f in root_folder.Folders:
                recurse(f)
                
        except Exception as e:
            print("Error fetching folder list: {}".format(e))
            
        return sorted(folders)

    def get_category_map(self):
        """Returns a dict of {CategoryName: HexColor}."""
        if not self.namespace: return {}
        
        mapping = {}
        try:
            if self.namespace.Categories.Count > 0:
                for cat in self.namespace.Categories:
                    c_enum = cat.Color
                    hex_code = self.OL_CAT_COLORS.get(c_enum, "#555555")
                    mapping[cat.Name] = hex_code
        except Exception as e:
             print("Error fetching categories: {}".format(e))
        return mapping

    def create_email(self):
        """Creates and displays a new email."""
        if not self.outlook: self.connect()
        try:
            mail = self.outlook.CreateItem(0) # olMailItem
            mail.Display()
            return True
        except Exception as e:
            print("Error creating email: {}".format(e))
            return False

    def create_appointment(self):
        """Creates and displays a new appointment."""
        if not self.outlook: self.connect()
        try:
            appt = self.outlook.CreateItem(1) # olAppointmentItem
            appt.Display()
            return True
        except Exception as e:
            print("Error creating appointment: {}".format(e))
            return False

    def create_meeting(self):
        """Creates and displays a new meeting."""
        if not self.outlook: self.connect()
        try:
            appt = self.outlook.CreateItem(1) # olAppointmentItem
            appt.MeetingStatus = 1 # olMeeting
            appt.Display()
            return True
        except Exception as e:
            print("Error creating meeting: {}".format(e))
            return False

    def create_task(self):
        """Creates and displays a new task."""
        if not self.outlook: self.connect()
        try:
            task = self.outlook.CreateItem(3) # olTaskItem
            task.Display()
            return True
        except Exception as e:
            print("Error creating task: {}".format(e))
            return False

    def get_due_status(self, account_names=None):
        """Checks for items due today or overdue. Returns dict {'tasks': status, 'calendar': status}."""
        status = {"tasks": None, "calendar": None}
        if not self.namespace: self.connect()
        try:
            now = datetime.now()
            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
            tomorrow = today + timedelta(days=1)
            
            # Formats for DASL queries
            today_str = today.strftime('%m/%d/%Y %I:%M %p')
            tom_str = tomorrow.strftime('%m/%d/%Y %I:%M %p')
            
            # Simple check across accounts
            for store in self._get_enabled_stores(account_names):
                try:
                    # Check Tasks
                    if status["tasks"] != "Overdue": # Optimization: Stop if worst case found
                        tasks = store.GetDefaultFolder(13)
                        base_filter = "[Complete] = False"
                        
                        # Overdue
                        overdue_filter = "{} AND [DueDate] < '{}'".format(base_filter, today_str)
                        try:
                            t_over = tasks.GetTable(overdue_filter)
                            if t_over.GetRowCount() > 0: status["tasks"] = "Overdue"
                        except: pass

                        if status["tasks"] is None:
                            # Today
                            today_filter = "{} AND [DueDate] >= '{}' AND [DueDate] < '{}'".format(base_filter, today_str, tom_str)
                            try:
                                t_today = tasks.GetTable(today_filter)
                                if t_today.GetRowCount() > 0: status["tasks"] = "Today"
                            except: pass
                    
                    # Check Appointments (Today)
                    if status["calendar"] is None:
                        cal = store.GetDefaultFolder(9)
                        cal_filter = "[Start] >= '{}' AND [Start] < '{}'".format(today_str, tom_str)
                        try:
                            t_cal = cal.GetTable(cal_filter)
                            if t_cal.GetRowCount() > 0: status["calendar"] = "Today"
                        except: pass

                except: continue
        except Exception as e:
            print("Error checking due items: {}".format(e))
        return status

    def send_email(self, subject, body, recipients):
        """
        Sends an email using the default account.
        recipients: string, semicolon separated
        """
        try:
             app = self._get_outlook_app()
             if not app: return False
             
             mail = app.CreateItem(0) # olMailItem
             mail.Subject = subject
             mail.Body = body
             mail.To = recipients
             mail.Send()
             return True
        except Exception as e:
             print("Error sending email: {}".format(e))
             return False



class FolderPickerFrame(tk.Frame):
    def __init__(self, parent, folders, callback, on_cancel, selected_paths=None):
        tk.Frame.__init__(self, parent)
        self.callback = callback
        self.on_cancel = on_cancel
        self.folders = folders
        self.selected_paths = set(selected_paths) if selected_paths else set()
        
        # Win11 Colors
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
                pil_img = pil_img.resize((20, 20), Image.Resampling.LANCZOS)
                self.close_icon = ImageTk.PhotoImage(pil_img)
                btn_close = tk.Label(header, image=self.close_icon, bg=self.colors["bg"], cursor="hand2")
             except:
                btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")
        else:
             btn_close = tk.Label(header, text="✕", bg=self.colors["bg"], fg="#CCCCCC", cursor="hand2")

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
            background="#2D2D30", 
            foreground="white", 
            fieldbackground="#2D2D30",
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
    def __init__(self, parent, folders, callback, selected_paths=None):
        tk.Toplevel.__init__(self, parent)
        self.callback = callback
        self.title("Select Folders")
        self.overrideredirect(True) 
        
        # Colors (Match settings panel / dark theme)
        self.config(bg="#202020", highlightbackground="#60CDFF", highlightthickness=1)
        
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

        self.picker = FolderPickerFrame(self, folders, on_done, on_cancel, selected_paths)
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

        # Help Text
        tk.Label(
            self, 
            text="Select folders to sync. Hold Shift/Ctrl for multiple.",
            bg=self.colors["bg"], fg="#888888", font=("Segoe UI", 9, "italic")
        ).pack(fill="x", padx=15, pady=(0, 10))
        
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
    def __init__(self, parent, accounts, current_enabled, callback):
        tk.Toplevel.__init__(self, parent)
        self.callback = callback
        self.colors = {
            "bg": "#202020", "fg": "#FFFFFF", "accent": "#60CDFF"
        }
        
        self.title("Enabled Accounts")
        self.overrideredirect(True)
        self.wm_attributes("-topmost", True)
        self.config(bg=self.colors["bg"])
        self.configure(highlightbackground=self.colors["accent"], highlightthickness=1)
        
        w, h = 450, 550
        x = parent.winfo_x() + 50
        y = parent.winfo_y() + 50
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
                 messagebox.showerror("Error", "Could not connect to Outlook Sidebar.")
                 return

             folders = sidebar.outlook_client.get_folder_list(account_name)
             
             if not folders:
                 messagebox.showwarning("No Folders", "Could not retrieve folder list for '{}'.".format(account_name))
                 return
                 
             FolderPickerWindow(self, folders, on_selected, selected_paths)
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

class FeedbackDialog(tk.Toplevel):
    def __init__(self, parent, outlook_client):
        tk.Toplevel.__init__(self, parent)
        self.outlook_client = outlook_client
        self.title("Send Feedback")
        self.overrideredirect(True) 
        self.geometry("400x320")
        
        # Make Modal
        self.transient(parent)
        self.grab_set()

        # Center
        try:
             px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 200
             py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 160
             self.geometry("+{}+{}".format(px, py))
        except: pass

        # Dark Theme Colors
        self.colors = {
            "bg_root": "#202020",
            "bg_card": "#2D2D30",
            "accent": "#60CDFF",
            "fg_text": "#FFFFFF",
            "fg_dim": "#AAAAAA"
        }
        self.configure(bg=self.colors["bg_root"], highlightthickness=1, highlightbackground=self.colors["accent"])

        # Title Bar (for dragging)
        title_bar = tk.Frame(self, bg=self.colors["bg_root"], height=30)
        title_bar.pack(fill="x", side="top")
        title_bar.bind("<Button-1>", self.start_move)
        title_bar.bind("<B1-Motion>", self.on_move)
        
        tk.Label(title_bar, text="Send Feedback", fg=self.colors["fg_text"], bg=self.colors["bg_root"], font=("Segoe UI", 10, "bold")).pack(side="left", padx=10)
        
        # Content
        content = tk.Frame(self, bg=self.colors["bg_root"], padx=15, pady=10)
        content.pack(fill="both", expand=True)
        
        tk.Label(content, text="Please describe your feedback, bug report, or suggestion:", 
                 fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 9), anchor="w").pack(fill="x", pady=(0, 5))
        
        self.txt_input = tk.Text(content, bg=self.colors["bg_card"], fg=self.colors["fg_text"], 
                                 insertbackground="white", font=("Segoe UI", 10), height=8, relief="flat", padx=5, pady=5)
        self.txt_input.pack(fill="both", expand=True, pady=5)
        self.txt_input.focus_set()

        # Buttons
        btn_frame = tk.Frame(content, bg=self.colors["bg_root"])
        btn_frame.pack(fill="x", pady=10)
        
        def on_send():
            text = self.txt_input.get("1.0", "end-1c").strip()
            if not text:
                messagebox.showwarning("Empty", "Please enter some feedback first.")
                return
            
            # Send
            # Recipient: Ralph (Developer)
            recipient = "ralph.marsh@coveya.co.uk"
            subject = "Outlook Sidebar Feedback ({})".format(VERSION)
            
            success = self.outlook_client.send_email(subject, text, recipient)
            if success:
                messagebox.showinfo("Sent", "Feedback sent successfully!")
                self.destroy()
            else:
                messagebox.showerror("Error", "Failed to send feedback.\nPlease check Outlook is running.")

        btn_cancel = tk.Label(btn_frame, text="Cancel", fg="#AAAAAA", bg=self.colors["bg_root"], 
                              font=("Segoe UI", 10), cursor="hand2", padx=15, pady=5)
        btn_cancel.pack(side="right", padx=5)
        btn_cancel.bind("<Button-1>", lambda e: self.destroy())
        
        btn_send = tk.Label(btn_frame, text="Send Feedback", fg="white", bg=self.colors["accent"], 
                            font=("Segoe UI", 10, "bold"), cursor="hand2", padx=15, pady=5)
        btn_send.pack(side="right", padx=5)
        btn_send.bind("<Button-1>", lambda e: on_send())
        
        # Hover effects
        def on_enter(w, bg): w.config(bg=bg)
        def on_leave(w, bg): w.config(bg=bg)
        
        btn_send.bind("<Enter>", lambda e: on_enter(btn_send, "#40b0ff"))
        btn_send.bind("<Leave>", lambda e: on_leave(btn_send, self.colors["accent"]))
        
        btn_cancel.bind("<Enter>", lambda e: btn_cancel.config(fg="white"))
        btn_cancel.bind("<Leave>", lambda e: btn_cancel.config(fg="#AAAAAA"))

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def on_move(self, event):
        x = self.winfo_x() + (event.x_root - self.x)
        y = self.winfo_y() + (event.y_root - self.y)
        self.geometry("+{}+{}".format(x, y))


class SettingsPanel(tk.Frame):
    """Inline settings panel that extends from the sidebar."""
    def __init__(self, parent, main_window, callback):
        tk.Frame.__init__(self, parent, bg="#202020")
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
        self.option_add('*TCombobox*Listbox.font', ("Segoe UI", 10))


        # Configure ttk.Style for comboboxes to fix dark theme visibility
        style = ttk.Style()
        style.theme_use('clam')  # Use clam theme for better dark mode support
        
        # Create dedicated style for Font Size combobox (not affecting other comboboxes)
        style.configure('FontSize.TCombobox',
            fieldbackground='#2d2d2d',  # Dark background for input field
            background='#2d2d2d',        # Dark background for button
            foreground='white',          # White text for selected value
            arrowcolor='white',          # White dropdown arrow
            bordercolor='#555555',       # Border color
            lightcolor='#2d2d2d',
            darkcolor='#2d2d2d',
            selectbackground='#007ACC',  # Blue highlight
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
        self.scroll_frame = ScrollableFrame(self, bg=self.colors["bg_root"])
        self.scroll_frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        main_content = self.scroll_frame.scrollable_frame
        main_content.config(bg=self.colors["bg_root"])

        

            
        # --- Sidebar Placement Setting REMOVED (Auto-snap implemented) ---
        # placement_frame = tk.Frame(self, bg=self.colors["bg_root"])
        # ...

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
                                 bg=self.colors["bg_card"], fg="white", bd=0, font=("Segoe UI", 10),
                                 highlightthickness=1, highlightbackground="#444444", pady=8)
        btn_accounts.pack(fill="x", padx=(18, 30), pady=(5, 5))

        # --- Email List Settings ---
        list_settings_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        list_settings_frame.pack(fill="x", padx=(18, 30), pady=(10, 0))
        
        print("DEBUG: Creating email settings checkboxes")
        self.show_read_var = tk.BooleanVar(value=self.main_window.show_read)
        print("DEBUG: show_read_var created with value: {}".format(self.show_read_var.get()))
        
        # Add trace callback
        def on_show_read_change(*args):
            print("DEBUG: show_read_var changed!")
            self.update_email_filters()
        
        self.show_read_var.trace("w", on_show_read_change)
        print("DEBUG: Trace added to show_read_var")
        
        self.chk_show_read = tk.Checkbutton(
            list_settings_frame, text="Include read email", 
            variable=self.show_read_var,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 10)
        )
        self.chk_show_read.grid(row=0, column=0, sticky="w", pady=(0, 5))
        print("DEBUG: show_read checkbox created and gridded")

        self.show_has_attachment_var = tk.BooleanVar(value=self.main_window.show_has_attachment)
        print("DEBUG: show_has_attachment_var created with value: {}".format(self.show_has_attachment_var.get()))
        
        # Add trace callback
        def on_show_attachment_change(*args):
            print("DEBUG: show_has_attachment_var changed!")
            self.update_email_filters()
        
        self.show_has_attachment_var.trace("w", on_show_attachment_change)
        print("DEBUG: Trace added to show_has_attachment_var")
        
        self.chk_has_attachment = tk.Checkbutton(
            list_settings_frame, text="Show if has Attachment", 
            variable=self.show_has_attachment_var,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 10)
        )
        self.chk_has_attachment.grid(row=1, column=0, sticky="w", pady=(0, 5))
        print("DEBUG: show_has_attachment checkbox created and gridded")
        
        # --- Email Window Content ---
        self.email_content_visible = False
        self.btn_email_content = tk.Button(
            list_settings_frame, text="Email Window Content",
            command=self.toggle_email_content_options,
            bg=self.colors["bg_card"], fg="white",
            font=("Segoe UI", 10),
            bd=0, padx=10, pady=4,
            cursor="hand2"
        )
        self.btn_email_content.grid(row=2, column=0, sticky="w", pady=(10, 5))
        
        self.email_content_frame = tk.Frame(list_settings_frame, bg=self.colors["bg_root"])
        self.email_content_frame.grid(row=3, column=0, sticky="w", padx=(20, 0))
        self.email_content_frame.grid_remove() # Initially hidden
        
        # Checkboxes
        self.email_show_sender_var = tk.BooleanVar(value=self.main_window.email_show_sender)
        tk.Checkbutton(self.email_content_frame, text="Who From", variable=self.email_show_sender_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg="white", 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground="white", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
                       
        self.email_show_subject_var = tk.BooleanVar(value=self.main_window.email_show_subject)
        tk.Checkbutton(self.email_content_frame, text="Subject Line", variable=self.email_show_subject_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg="white", 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground="white", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w")

        self.email_show_body_var = tk.BooleanVar(value=self.main_window.email_show_body)
        tk.Checkbutton(self.email_content_frame, text="Content Body", variable=self.email_show_body_var, 
                       command=self.update_email_filters, bg=self.colors["bg_root"], fg="white", 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground="white", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w")
        
        # Move "Show Content on Hover" OUT of here
        
        # Number of Lines Selector
        lines_frame = tk.Frame(self.email_content_frame, bg=self.colors["bg_root"])
        lines_frame.grid(row=4, column=0, sticky="w", pady=(5,0))
        tk.Label(lines_frame, text="Lines:", bg=self.colors["bg_root"], fg="#cccccc", font=("Segoe UI", 10)).pack(side="left")
        
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

        # "Show Content on Hover" - Independent Setting (Below the content frame)
        self.show_hover_content_var = tk.BooleanVar(value=self.main_window.show_hover_content)
        chk_hover = tk.Checkbutton(
             list_settings_frame, text="Show Content on Hover", 
             variable=self.show_hover_content_var, 
             command=self.update_email_filters, 
             bg=self.colors["bg_root"], fg="white", 
             selectcolor=self.colors["bg_card"], 
             activebackground=self.colors["bg_root"], 
             activeforeground="white", 
             font=("Segoe UI", 10)
        )
        chk_hover.grid(row=4, column=0, sticky="w", pady=(5, 5))

        # --- Interaction Settings (Merged into Email Settings) ---
        interaction_frame = tk.Frame(main_content, bg=self.colors["bg_root"])
        interaction_frame.pack(fill="x", padx=(18, 30), pady=(5, 10))
        
        self.buttons_on_hover_var = tk.BooleanVar(value=self.main_window.buttons_on_hover)
        tk.Checkbutton(interaction_frame, text="Show Buttons on Hover", variable=self.buttons_on_hover_var, 
                       command=self.update_interaction_settings, bg=self.colors["bg_root"], fg="white", 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground="white", font=("Segoe UI", 9)).pack(side="left")
                       
        self.email_double_click_var = tk.BooleanVar(value=self.main_window.email_double_click)
        tk.Checkbutton(interaction_frame, text="Double Click to Open", variable=self.email_double_click_var, 
                       command=self.update_interaction_settings, bg=self.colors["bg_root"], fg="white", 
                       selectcolor=self.colors["bg_card"], activebackground=self.colors["bg_root"], 
                       activeforeground="white", font=("Segoe UI", 9)).pack(side="left", padx=10)

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
                bg=self.colors["bg_root"], fg="white",
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground="white",
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
            btn_pick = tk.Label(f_frame, text="...", bg=self.colors["bg_card"], fg="white", font=("Segoe UI", 8), width=3, cursor="hand2")
            btn_pick.pack(side="left", padx=(5,0), fill="y")
            
            # Bind picker
            def open_picker(event, entry=e_folder):
                folders = self.main_window.outlook_client.get_folder_list()
                FolderPickerWindow(self, folders if folders else ["Inbox"], 
                                   lambda path: (entry.delete(0, tk.END), entry.insert(0, path), self.update_button_config()))

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
        # Removed Hover Leave binding to keep open until clicked


        # Toggle button (Arrow) inside container
        self.btn_toggle_followup = tk.Label(
            self.followup_container, text="▼",
            bg=self.colors["bg_root"], fg="#AAAAAA",
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_followup.pack(side="top", anchor="w", pady=(2, 0))
        self.btn_toggle_followup.bind("<Button-1>", lambda e: self.toggle_followup_visibility())
        
        # Initially hide button if Follow-up Flags is unchecked
        if not self.main_window.reminder_show_flagged:
            self.followup_container.grid_remove() # Hide entire container
        
        # Container for due date checkboxes (conditionally shown)
        # MOVED inside unified container
        self.followup_options_frame = tk.Frame(self.followup_container, bg=self.colors["bg_root"])
        # self.followup_options_frame.pack(side="top", anchor="w", padx=(0, 0)) # Start Hidden

        
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
                bg=self.colors["bg_root"], fg="white",
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground="white",
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
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
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
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9, "bold")
        )
        chk_categorized.grid(row=3, column=0, sticky="w", pady=(0, 5), columnspan=3)
        
        # --- 3. Importance ---
        self.reminder_show_importance_var = tk.BooleanVar(value=self.main_window.reminder_show_importance)  # Initialize from config
        chk_importance = tk.Checkbutton(
            reminder_frame, text="Importance", 
            variable=self.reminder_show_importance_var,
            command=self.toggle_importance_options,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9, "bold")
        )
        chk_importance.grid(row=4, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button for showing/hiding options
        self.importance_options_visible = False
        
        # Unified Container for Hover Logic
        # Unified Container for Hover Logic
        self.importance_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.importance_container.grid(row=4, column=1, sticky="nw", rowspan=2, padx=(5, 0))
        # Removed leave binding

        self.btn_toggle_importance = tk.Label(
            self.importance_container, text="▼",
            bg=self.colors["bg_root"], fg="#AAAAAA",
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_importance.pack(side="top", anchor="w", pady=(2, 0))
        self.btn_toggle_importance.bind("<Button-1>", lambda e: self.toggle_importance_visibility())
        self.btn_toggle_importance.grid_remove()  # Initially hidden
        
        # Container for importance checkboxes
        # MOVED inside unified container
        self.importance_options_frame = tk.Frame(self.importance_container, bg=self.colors["bg_root"])
        # self.importance_options_frame.pack(side="top", anchor="w", padx=(0, 0)) # Start Hidden

        
        # Adjust master checkbox alignment
        chk_importance.grid(row=4, column=0, sticky="nw", pady=(0, 5))
        
        self.reminder_high_importance_var = tk.BooleanVar(value=self.main_window.reminder_high_importance)
        chk_high = tk.Checkbutton(
            self.importance_options_frame, text="High", 
            variable=self.reminder_high_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        # Use grid inside packed frame
        chk_high.grid(row=0, column=0, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_normal_importance_var = tk.BooleanVar(value=self.main_window.reminder_normal_importance)
        chk_normal = tk.Checkbutton(
            self.importance_options_frame, text="Normal", 
            variable=self.reminder_normal_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_normal.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_low_importance_var = tk.BooleanVar(value=self.main_window.reminder_low_importance)
        chk_low = tk.Checkbutton(
            self.importance_options_frame, text="Low", 
            variable=self.reminder_low_importance_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_low.grid(row=0, column=2, sticky="w", pady=2)
        
        # Initially hide
        self.importance_options_frame.grid_remove()
        
        # --- 4. Meetings ---
        self.reminder_show_meetings_var = tk.BooleanVar(value=self.main_window.reminder_show_meetings)  # Initialize from config
        chk_meetings = tk.Checkbutton(
            reminder_frame, text="Meetings", 
            variable=self.reminder_show_meetings_var,
            command=self.toggle_meetings_options,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9, "bold")
        )
        chk_meetings.grid(row=6, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button (Arrow)
        self.meetings_options_visible = False
        
        # Unified Container
        self.meetings_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.meetings_container.grid(row=6, column=1, sticky="nw", rowspan=2, padx=(5, 0))
        # Removed leave binding

        self.btn_toggle_meetings = tk.Label(
            self.meetings_container, text="▼",
            bg=self.colors["bg_root"], fg="#AAAAAA",
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_meetings.pack(side="top", anchor="w", pady=(2, 0))
        self.btn_toggle_meetings.bind("<Button-1>", lambda e: self.toggle_meetings_visibility())
        self.btn_toggle_meetings.grid_remove()  # Initially hidden
        
        # Track if defaults have been applied
        self.meetings_defaults_applied = False
        
        # Container for meeting options
        # MOVED inside unified container
        self.meetings_options_frame = tk.Frame(self.meetings_container, bg=self.colors["bg_root"])
        # self.meetings_options_frame.pack(side="top", anchor="w", padx=(0, 0)) # Start Hidden

        
        # Adjust master checkbox alignment
        chk_meetings.grid(row=6, column=0, sticky="nw", pady=(0, 5))
        
        # Status checkboxes
        self.reminder_pending_meetings_var = tk.BooleanVar(value=self.main_window.reminder_pending_meetings)
        chk_pending = tk.Checkbutton(
            self.meetings_options_frame, text="Pending", 
            variable=self.reminder_pending_meetings_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_pending.grid(row=0, column=0, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_accepted_meetings_var = tk.BooleanVar(value=self.main_window.reminder_accepted_meetings)
        chk_accepted = tk.Checkbutton(
            self.meetings_options_frame, text="Accepted", 
            variable=self.reminder_accepted_meetings_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_accepted.grid(row=0, column=1, sticky="w", pady=2)
        
        self.reminder_declined_meetings_var = tk.BooleanVar(value=self.main_window.reminder_declined_meetings)
        chk_rejected = tk.Checkbutton(
            self.meetings_options_frame, text="Rejected", 
            variable=self.reminder_declined_meetings_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_rejected.grid(row=1, column=0, sticky="w", pady=2)
        
        # Date filters
        # Options: ["Today", "Tomorrow", "Next 7 Days", "Next X Days"]
        self.meeting_date_options = ["Today", "Tomorrow", "Next 7 Days"]
        self.meeting_date_vars = {}
        
        # Grid layout for first 3
        for idx, option in enumerate(self.meeting_date_options):
            # Init based on main window state
            is_checked = option in self.main_window.reminder_meeting_dates
            var = tk.BooleanVar(value=is_checked)
            self.meeting_date_vars[option] = var
            
            chk = tk.Checkbutton(
                self.meetings_options_frame, text=option,
                variable=var,
                command=self.update_reminder_filters,
                bg=self.colors["bg_root"], fg="white",
                selectcolor=self.colors["bg_card"],
                activebackground=self.colors["bg_root"],
                activeforeground="white",
                font=("Segoe UI", 9)
            )
            # Row 2: Today, Tomorrow. Row 3: Next 7 Days
            r = 2 + (idx // 2)
            c = idx % 2
            chk.grid(row=r, column=c, sticky="w", padx=(0, 15), pady=2)

        # "Next X Days" with Entry
        is_custom = "Custom" in self.main_window.reminder_meeting_dates
        self.meeting_date_vars["Custom"] = tk.BooleanVar(value=is_custom)
        self.reminder_custom_days_var = tk.StringVar(value=str(getattr(self.main_window, "reminder_custom_days", 30)))
        
        f_custom = tk.Frame(self.meetings_options_frame, bg=self.colors["bg_root"])
        f_custom.grid(row=3, column=1, sticky="w", padx=0, pady=2)
        
        chk_custom = tk.Checkbutton(
            f_custom, text="Next",
            variable=self.meeting_date_vars["Custom"],
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_custom.pack(side="left")
        
        spin_days = tk.Spinbox(
            f_custom, from_=1, to=365,
            width=3,
            textvariable=self.reminder_custom_days_var,
            font=("Segoe UI", 9),
            bg=self.colors["input_bg"],
            fg="white",
            buttonbackground=self.colors["bg_card"]
        )
        spin_days.pack(side="left", padx=2)
        # trace changes to update immediately
        self.reminder_custom_days_var.trace("w", lambda *a: self.update_reminder_filters())
        lbl_days = tk.Label(f_custom, text="Days", bg=self.colors["bg_root"], fg="white", font=("Segoe UI", 9))
        lbl_days.pack(side="left")

        # ----------------------------------------------------
        # 8. Feedback (Moved from Help Panel)
        # ----------------------------------------------------
        self.create_section_header(scroll_content, "Feedback", pady=(20, 5))
        
        f_feedback = tk.Frame(scroll_content, bg=self.colors["bg_root"])
        f_feedback.pack(fill="x", anchor="w", padx=20, pady=5)
        
        tk.Label(f_feedback, text="Found a bug or have an idea?", 
                 fg="#DDDDDD", bg=self.colors["bg_root"], font=("Segoe UI", 9)).pack(anchor="w")
                 
        btn_fb = tk.Label(f_feedback, text="Send Feedback", fg="white", bg=self.colors["accent"], 
                          font=("Segoe UI", 9, "bold"), cursor="hand2", padx=15, pady=6)
        btn_fb.pack(anchor="w", pady=(5, 0))
        
        def open_fb(e):
            FeedbackDialog(self.main_window, self.main_window.outlook_client)
            
        btn_fb.bind("<Button-1>", open_fb)
        btn_fb.bind("<Enter>", lambda e: btn_fb.config(bg="#40b0ff"))
        btn_fb.bind("<Leave>", lambda e: btn_fb.config(bg=self.colors["accent"]))
        
        # Initially hide
        self.meetings_options_frame.grid_remove()
        
        # --- 5. Tasks ---
        self.reminder_show_tasks_var = tk.BooleanVar(value=self.main_window.reminder_show_tasks)  # Initialize from config
        chk_tasks_master = tk.Checkbutton(
            reminder_frame, text="Tasks", 
            variable=self.reminder_show_tasks_var,
            command=self.toggle_tasks_options,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9, "bold")
        )
        chk_tasks_master.grid(row=8, column=0, sticky="w", pady=(0, 5))
        
        # Toggle button (Arrow)
        self.tasks_options_visible = False
        
        # Unified Container
        self.tasks_container = tk.Frame(reminder_frame, bg=self.colors["bg_root"])
        self.tasks_container.grid(row=8, column=1, sticky="nw", rowspan=2, padx=(5, 0))
        # Removed leave binding

        self.btn_toggle_tasks = tk.Label(
            self.tasks_container, text="▼",
            bg=self.colors["bg_root"], fg="#AAAAAA",
            font=("Segoe UI", 8),
            cursor="hand2"
        )
        self.btn_toggle_tasks.pack(side="top", anchor="w", pady=(2, 0))
        self.btn_toggle_tasks.bind("<Button-1>", lambda e: self.toggle_tasks_visibility())
        self.btn_toggle_tasks.grid_remove()  # Initially hidden
        
        # Container for task options
        # MOVED inside unified container
        self.tasks_options_frame = tk.Frame(self.tasks_container, bg=self.colors["bg_root"])
        # self.tasks_options_frame.pack(side="top", anchor="w", padx=(0, 0)) # Start Hidden


        # Adjust master checkbox alignment
        chk_tasks_master.grid(row=8, column=0, sticky="nw", pady=(0, 5))
        
        self.reminder_tasks_var = tk.BooleanVar(value=self.main_window.reminder_tasks)
        chk_tasks = tk.Checkbutton(
            self.tasks_options_frame, text="Tasks", 
            variable=self.reminder_tasks_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_tasks.grid(row=0, column=0, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_todo_var = tk.BooleanVar(value=self.main_window.reminder_todo)
        chk_todo = tk.Checkbutton(
            self.tasks_options_frame, text="To-Do", 
            variable=self.reminder_todo_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_todo.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=2)
        
        self.reminder_has_reminder_var = tk.BooleanVar(value=self.main_window.reminder_has_reminder)
        chk_has_reminder = tk.Checkbutton(
            self.tasks_options_frame, text="Has Reminder", 
            variable=self.reminder_has_reminder_var,
            command=self.update_reminder_filters,
            bg=self.colors["bg_root"], fg="white",
            selectcolor=self.colors["bg_card"],
            activebackground=self.colors["bg_root"],
            activeforeground="white",
            font=("Segoe UI", 9)
        )
        chk_has_reminder.grid(row=0, column=2, sticky="w", pady=2)
        
        # Initially hide
        self.tasks_options_frame.grid_remove()



        # --- Icon Brightness Setting REMOVED ---
        # User requested fixed 75% brightness, slider removed.
        
        # Version Label REMOVED


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

    def toggle_email_content_options(self):
        """Show/hide email content options."""
        if self.email_content_visible:
            self.email_content_frame.grid_remove()
            self.email_content_visible = False
        else:
            self.email_content_frame.grid()
            self.email_content_visible = True

    def update_email_filters(self, event=None):
        """Apply email filter changes immediately."""
        print("DEBUG: update_email_filters called")
        self.main_window.show_read = self.show_read_var.get()
        self.main_window.show_has_attachment = self.show_has_attachment_var.get()
        
        # Email Window Content
        self.main_window.email_show_sender = self.email_show_sender_var.get()
        self.main_window.email_show_subject = self.email_show_subject_var.get()
        self.main_window.email_show_body = self.email_show_body_var.get()
        self.main_window.show_hover_content = self.show_hover_content_var.get() # Persist Hover Setting
        try:
            self.main_window.email_body_lines = int(self.email_body_lines_var.get())
        except:
            self.main_window.email_body_lines = 1
            
        print("DEBUG: Content Settings: Sender={}, Subject={}, Body={}, Lines={}".format(self.main_window.email_show_sender, self.main_window.email_show_subject, self.main_window.email_show_body, self.main_window.email_body_lines))
        
        self.main_window.save_config()
        print("DEBUG: Calling callback: {}".format(self.callback))
        self.callback()  # refresh_emails
        print("DEBUG: Callback completed")

    def toggle_followup_options(self):
        """Show/hide follow-up due date options based on checkbox state."""
        if self.reminder_show_flagged_var.get():
            # Set default selections (Today, Tomorrow, Overdue) on first check
            if not getattr(self, "followup_defaults_applied", False):
                if "Today" in self.due_vars: self.due_vars["Today"].set(True)
                if "Tomorrow" in self.due_vars: self.due_vars["Tomorrow"].set(True)
                if "Overdue" in self.due_vars: self.due_vars["Overdue"].set(True)
                self.followup_defaults_applied = True
            
            self.followup_options_frame.grid()
            self.btn_toggle_followup.grid()  # Show the toggle button
        else:
            self.followup_options_frame.grid_remove()
            self.btn_toggle_followup.grid_remove()  # Hide the toggle button
        self.update_reminder_filters()

    def close_all_sections(self, except_section=None):
        """Accordion Logic: Close all reminder sections except the one being opened."""
        sections = [
            ("followup", self.followup_options_visible, self.toggle_followup_visibility),
            ("importance", self.importance_options_visible, self.toggle_importance_visibility),
            ("meetings", self.meetings_options_visible, self.toggle_meetings_visibility),
            ("tasks", self.tasks_options_visible, self.toggle_tasks_visibility)
        ]
        
        for name, is_visible, toggle_func in sections:
            if is_visible and name != except_section:
                toggle_func(force_hide=True)

    def toggle_followup_visibility(self, force_hide=False):
        """Toggle visibility of follow-up options and update button text."""
        if self.followup_options_visible or force_hide:
            # Hide options
            self.followup_options_frame.pack_forget()
            self.btn_toggle_followup.config(text="▼")
            self.followup_options_visible = False
        else:
            # Accordion: Close others
            self.close_all_sections("followup")
            # Show options
            self.followup_options_frame.pack(side="top", anchor="w")
            self.btn_toggle_followup.config(text="▲")
            self.followup_options_visible = True

    def toggle_all_due_options(self):
        """Select or deselect all due date options."""
        all_checked = self.due_all_var.get()
        for var in self.due_vars.values():
            var.set(all_checked)
        self.update_reminder_filters()

    def toggle_importance_options(self):
        """Show/hide importance options based on checkbox state."""
        if self.reminder_show_importance_var.get():
            # Set default selections (all) if first time checking
            if not self.reminder_high_importance_var.get() and not self.reminder_normal_importance_var.get() and not self.reminder_low_importance_var.get():
                self.reminder_high_importance_var.set(True)
                self.reminder_normal_importance_var.set(True)
                self.reminder_low_importance_var.set(True)
            self.importance_options_frame.grid()
            self.btn_toggle_importance.grid()
        else:
            self.importance_options_frame.grid_remove()
            self.btn_toggle_importance.grid_remove()
        self.update_reminder_filters()

    def toggle_importance_visibility(self, force_hide=False):
        """Toggle visibility of importance options and update button text."""
        if self.importance_options_visible or force_hide:
            self.importance_options_frame.pack_forget()
            self.btn_toggle_importance.config(text="▼")
            self.importance_options_visible = False
        else:
            self.close_all_sections("importance")
            self.importance_options_frame.pack(side="top", anchor="w")
            self.btn_toggle_importance.config(text="▲")
            self.importance_options_visible = True

    def toggle_meetings_options(self):
        """Show/hide meetings options based on checkbox state."""
        if self.reminder_show_meetings_var.get():
            # Set default selections (Pending, Accepted, Declined, Today, Tomorrow) on first enable
            if not self.meetings_defaults_applied:
                self.reminder_pending_meetings_var.set(True)
                self.reminder_accepted_meetings_var.set(True)
                self.reminder_declined_meetings_var.set(True)
                # Set date filters: Today and Tomorrow
                if "Today" in self.meeting_date_vars:
                    self.meeting_date_vars["Today"].set(True)
                if "Tomorrow" in self.meeting_date_vars:
                    self.meeting_date_vars["Tomorrow"].set(True)
                self.meetings_defaults_applied = True
            self.meetings_options_frame.grid()
            self.btn_toggle_meetings.grid()
        else:
            self.meetings_options_frame.grid_remove()
            self.btn_toggle_meetings.grid_remove()
        self.update_reminder_filters()

    def toggle_meetings_visibility(self, force_hide=False):
        """Toggle visibility of meetings options and update button text."""
        if self.meetings_options_visible or force_hide:
            self.meetings_options_frame.pack_forget()
            self.btn_toggle_meetings.config(text="▼")
            self.meetings_options_visible = False
        else:
            self.close_all_sections("meetings")
            self.meetings_options_frame.pack(side="top", anchor="w")
            self.btn_toggle_meetings.config(text="▲")
            self.meetings_options_visible = True

    def toggle_tasks_options(self):
        """Show/hide tasks options based on checkbox state."""
        if self.reminder_show_tasks_var.get():
            self.tasks_options_frame.grid()
            self.btn_toggle_tasks.grid()
        else:
            self.tasks_options_frame.grid_remove()
            self.btn_toggle_tasks.grid_remove()
        self.update_reminder_filters()

    def toggle_tasks_visibility(self, force_hide=False):
        """Toggle visibility of tasks options and update button text."""
        if self.tasks_options_visible or force_hide:
            self.tasks_options_frame.pack_forget()
            self.btn_toggle_tasks.config(text="▼")
            self.tasks_options_visible = False
        else:
            self.close_all_sections("tasks")
            self.tasks_options_frame.pack(side="top", anchor="w")
            self.btn_toggle_tasks.config(text="▲")
            self.tasks_options_visible = True


    def update_reminder_filters(self, event=None):
        """Apply reminder filter changes immediately."""
        self.main_window.reminder_show_flagged = self.reminder_show_flagged_var.get()
        
        # Collect selected due date filters from checkboxes
        selected_due_filters = [option for option, var in self.due_vars.items() if var.get()]
        self.main_window.reminder_due_filters = selected_due_filters  # Store as list
        
        self.main_window.reminder_show_categorized = self.reminder_show_categorized_var.get()
        self.main_window.reminder_show_importance = self.reminder_show_importance_var.get()
        self.main_window.reminder_high_importance = self.reminder_high_importance_var.get()
        self.main_window.reminder_normal_importance = self.reminder_normal_importance_var.get()
        self.main_window.reminder_low_importance = self.reminder_low_importance_var.get()
        self.main_window.reminder_show_meetings = self.reminder_show_meetings_var.get()
        self.main_window.reminder_pending_meetings = self.reminder_pending_meetings_var.get()
        self.main_window.reminder_accepted_meetings = self.reminder_accepted_meetings_var.get()
        self.main_window.reminder_declined_meetings = self.reminder_declined_meetings_var.get()
        
        # Collect selected meeting date filters
        selected_meeting_dates = [option for option, var in self.meeting_date_vars.items() if var.get()]
        # Collect selected meeting date filters
        selected_meeting_dates = [option for option, var in self.meeting_date_vars.items() if var.get()]
        self.main_window.reminder_meeting_dates = selected_meeting_dates
        
        try:
             days = int(self.reminder_custom_days_var.get())
             self.main_window.reminder_custom_days = days
        except:
             pass
        
        self.main_window.reminder_show_tasks = self.reminder_show_tasks_var.get()
        self.main_window.reminder_tasks = self.reminder_tasks_var.get()
        self.main_window.reminder_todo = self.reminder_todo_var.get()
        self.main_window.reminder_has_reminder = self.reminder_has_reminder_var.get()
        self.main_window.reminder_has_reminder = self.reminder_has_reminder_var.get()
        self.main_window.reminder_has_reminder = self.reminder_has_reminder_var.get()
        self.main_window.save_config()
        self.main_window.refresh_reminders()

    def update_interaction_settings(self):
        """Allow checkbox command to update interaction settings only."""
        self.main_window.buttons_on_hover = self.buttons_on_hover_var.get()
        self.main_window.email_double_click = self.email_double_click_var.get()
        self.main_window.save_config()
        self.callback()
    
    def update_button_config(self):
        """Apply button config changes immediately (Original Logic)."""
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
        
        # Also update interaction settings
        self.update_interaction_settings()

    def update_button_action(self, idx, event=None):
        """Deprecated: Kept for compatibility if old bindings trigger."""
        self.update_button_config()

    def update_button_folder(self, idx, event=None):
         """Deprecated: Kept for compatibility if old bindings trigger."""
         self.update_button_config()

    def open_folder_picker(self, account_name):
        """Opens the inline folder picker for account settings."""
        try:
             if not self.main_window or not hasattr(self.main_window, "outlook_client"):
                 print("Error: Could not locate OutlookClient")
                 return

             folders = self.main_window.outlook_client.get_folder_list(account_name)
             
             def on_return():
                 if hasattr(self, "folder_picker"):
                     self.folder_picker.destroy()
                 self.scroll_frame.pack(fill="both", expand=True, padx=2, pady=2)

             def on_pick(paths):
                 self.vars[account_name]["email_folders"] = paths
                 
             if not folders:
                 messagebox.showwarning("No Folders", "Could not retrieve folder list for '{}'.".format(account_name))
                 return
                 
             self.scroll_frame.pack_forget()
             self.folder_picker = FolderPickerFrame(self, folders, on_pick, on_return)
             self.folder_picker.pack(fill="both", expand=True)
        except Exception as e:
            print("Error opening folder picker: {}".format(e))

    def browse_folder(self, idx):
        """Open folder picker for button at idx."""
        try:
            folders = self.main_window.outlook_client.get_folder_list()
            
            def on_select(paths):
                # User selected one or more folders. We need just one? Or path string?
                # Button config expects single path usually, or list.
                # Just take the first one or join them?
                # Legacy behavior: "path". 
                if paths:
                    self.btn_configs[idx]["folder_var"].set(paths[0])
                    self.update_button_folder(idx)
            
            # Since we are browsing for a setting, we can use a Toplevel wrapper for the Frame
            top = tk.Toplevel(self)
            top.title("Select Folder")
            top.geometry("350x450")
            top.config(bg="#202020")
            
            def on_cancel():
                top.destroy()
                
            picker = FolderPickerFrame(top, folders, on_select, on_cancel)
            picker.pack(fill="both", expand=True)
            
        except Exception as e:
            print("Error browsing folders: {}".format(e))

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
        
        # Apply window mode to main window layout
        self.main_window.window_mode = mode
        self.main_window.save_config()
        self.main_window.apply_window_layout()

    def close_panel(self):
        """Close the settings panel."""
        self.main_window.toggle_settings_panel()
        
class HelpPanel(tk.Frame):
    """Inline help panel that extends from the sidebar."""
    def __init__(self, parent, main_window):
        tk.Frame.__init__(self, parent, bg="#202020")
        self.main_window = main_window
        
        # --- Windows 11 Dark Theme ---
        self.colors = {
            "bg_root": "#202020",       # Deep Dark
            "bg_card": "#2D2D30",       # Input BG
            "accent": "#60CDFF",        # Win11 Blue
            "fg_text": "#FFFFFF",
            "fg_dim": "#A0A0A0",
        }
        
        # Frame styling
        self.config(bg=self.colors["bg_root"])
        self.configure(highlightbackground="#444444", highlightthickness=1)
        
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
        btn_close = tk.Label(header, text="✕", fg="#FF4444", bg=self.colors["bg_root"], font=("Arial", 14, "bold"), cursor="hand2")
        btn_close.pack(side="right", padx=10)
        btn_close.bind("<Button-1>", lambda e: self.main_window.toggle_help_panel())

        # --- Scrollable Container ---
        self.scroll_frame = ScrollableFrame(self, bg=self.colors["bg_root"])
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
            tk.Label(frame, text=text, fg="#DDDDDD", bg=self.colors["bg_root"], 
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
            
        AccountSelectionDialog(self.winfo_toplevel(), accounts, self.main_window.enabled_accounts, on_save)

class SidebarWindow(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        # --- Configuration ---
        self.min_width = 300  
        self.hot_strip_width = 16 # Customized
        self.expanded_width = 300
        self.is_pinned = True
        self.is_expanded = False
        self.dock_side = "Left" # "Left" or "Right"
        self.font_family = "Segoe UI"
        self.font_size = 9
        self.poll_interval = 30 # seconds
        self.show_read = False
        self.show_has_attachment = False  # Filter for emails with attachments
        self.only_flagged = False
        self.include_read_flagged = True
        self.include_read_flagged = True
        self.flag_date_filter = "Anytime"
        
        # Email Content Settings
        self.email_show_sender = True
        self.email_show_subject = True
        self.email_show_body = False
        self.email_body_lines = 2
        
        self.hover_delay = 500 # ms
        self._hover_timer = None
        self._collapse_timer = None
        
        # Settings Panel State
        self.settings_panel_open = False
        self.settings_panel = None
        self.settings_panel_width = 370
        
        # Window Mode State
        self.window_mode = "dual"  # "single" (just emails) or "dual" (emails + reminder list)
        self.split_sash_pos = 0 # 0 means auto/default
        
        # Reminder Filter State
        self.reminder_show_flagged = True  # Default ON
        self.reminder_due_filters = ["No Date"]  # List of selected due date filters
        self.reminder_show_categorized = True
        self.reminder_categories = []  # List of selected categories
        
        # Importance
        self.reminder_show_importance = True # Master toggle
        self.reminder_high_importance = False
        self.reminder_normal_importance = False
        self.reminder_low_importance = False
        
        # Meetings
        self.reminder_show_meetings = True # Master toggle
        self.reminder_pending_meetings = True
        self.reminder_accepted_meetings = True # Default ON
        self.reminder_declined_meetings = True # Changed to True as per user intent "Might want to know"
        self.reminder_meeting_dates = ["Today", "Tomorrow"]
        self.reminder_custom_days = 30
        
        # Tasks
        self.reminder_show_tasks = True # Master toggle
        self.reminder_tasks = True
        self.reminder_todo = True
        self.reminder_has_reminder = True
        
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
        self.buttons_on_hover = False
        self.buttons_on_hover = True
        self.email_double_click = True
        
        self.help_panel = None
        self.help_panel_open = False
        
        # Account Settings
        self.enabled_accounts = {} # {"Name": {"email": True, "calendar": True}}
        
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
        # Find the true top-level window handle (root owner)
        self.hwnd = self.winfo_id()
        temp_hwnd = self.hwnd
        while True:
            parent = ctypes.windll.user32.GetParent(temp_hwnd)
            if not parent:
                break
            temp_hwnd = parent
        self.hwnd = temp_hwnd

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
        # 1. Outlook Button (Rightmost)
        if os.path.exists("icon2/email.png"):
             # Email: 32x32 (Color: White)
             try:
                img = self.load_icon_colored("icon2/email.png", size=(32, 32), color="#FFFFFF")
                self.image_cache["outlook_footer"] = img
                self.btn_outlook = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                self.btn_outlook.pack(side="right", padx=(5, 10), pady=5)
                self.btn_outlook.bind("<Button-1>", lambda e: self.open_outlook_app())
                ToolTip(self.btn_outlook, "Open Outlook")
             except Exception as e:
                print("Error loading Outlook icon: {}".format(e))

        # 0. Close Button (Leftmost)
        # Use existing icon logic
        if os.path.exists("icon2/close-window.png"):
             try:
                # Close: 30x30, Red (#E81123 is standard Windows close red, or just Red)
                img = self.load_icon_colored("icon2/close-window.png", size=(30, 30), color="#FF4444")
                if img:
                    self.image_cache["close_footer"] = img
                    self.btn_close = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                else: 
                     raise Exception("Load failed")
             except Exception as e:
                print("Error loading Close icon: {}".format(e))
                self.btn_close = tk.Label(self.footer, text="✕", bg="#444444", fg="#aaaaaa", font=("Arial", 12), cursor="hand2")
        else:
            self.btn_close = tk.Label(self.footer, text="✕", bg="#444444", fg="#aaaaaa", font=("Arial", 12), cursor="hand2")
            
        self.btn_close.pack(side="left", padx=10, pady=5)
        self.btn_close.bind("<Button-1>", lambda e: self.quit_application())
        ToolTip(self.btn_close, "Close Application")
        
        # Version Label
        self.lbl_version = tk.Label(self.footer, text=VERSION, bg="#444444", fg="#888888", font=("Segoe UI", 8))
        self.lbl_version.pack(side="left", padx=5, pady=5)
        ToolTip(self.lbl_version, "App Version: {}".format(VERSION))
                 
        # 2. Calendar Button (Next to Outlook)
        # 2. Calendar Button (Next to Outlook)
        if os.path.exists("icon2/calendar.png"):
             # Calendar: Reduced to 28x28 (Color: White)
             try:
                img = self.load_icon_colored("icon2/calendar.png", size=(28, 28), color="#FFFFFF")
                self.image_cache["calendar_footer"] = img
                self.btn_calendar = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                self.btn_calendar.pack(side="right", padx=5, pady=5)
                self.btn_calendar.bind("<Button-1>", lambda e: self.open_calendar_app())
                ToolTip(self.btn_calendar, "Open Calendar")
             except Exception as e:
                print("Error loading Calendar icon: {}".format(e))

        # Quick Create Button (Plus)
        if os.path.exists("icon2/plus.png"):
             try:
                # Initial placeholder - update_quick_create_icon will set the correct one
                img = self.load_icon_colored("icon2/plus.png", size=(26, 26), color="#555555") 
                self.image_cache["quick_create"] = img
                self.btn_quick_create = tk.Label(self.footer, image=img, bg="#444444", cursor="hand2")
                self.btn_quick_create.pack(side="right", padx=5, pady=5)
                self.btn_quick_create.bind("<Button-1>", lambda e: self.handle_quick_create())
                ToolTip(self.btn_quick_create, "Quick Create")
                
                # Apply initial state
                self.update_quick_create_icon()
             except Exception as e:
                print("Error loading Quick Create icon: {}".format(e))

        # Header
        self.header = tk.Frame(self.main_frame, bg="#444444", height=40)
        self.header.pack(fill="x", side="top")
        
        # Header Dragging
        self.header.bind("<Button-1>", self.start_window_drag)
        self.header.bind("<B1-Motion>", self.on_window_drag)
        self.header.bind("<ButtonRelease-1>", self.stop_window_drag)
        
        # Title
        # Title
        self.lbl_title = tk.Label(self.header, text="InboxBar", bg="#444444", fg="white", font=(self.font_family, 10, "bold"))
        self.lbl_title.pack(side="left", padx=10)
        self.lbl_title.bind("<Button-1>", self.start_window_drag)
        self.lbl_title.bind("<B1-Motion>", self.on_window_drag)
        self.lbl_title.bind("<ButtonRelease-1>", self.stop_window_drag)

        # Pin Button / Logo (Custom Canvas)
        # Pin Button / Logo
        if os.path.exists("icon2/pin1.png"):
             try:
                 # Pin: 24x24 (Leave). Use pin1 for both.
                 # Active (Pinned) = White (#FFFFFF)
                 # Inactive (Unpinned) = Grey (#888888)
                 
                 # 1. Active State (Pinned) = White
                 self.icon_pin_active = self.load_icon_colored("icon2/pin1.png", size=(24, 24), color="#FFFFFF")
                 
                 # 2. Inactive State (Unpinned) = Black (Natural)
                 self.icon_pin_inactive = self.load_icon_colored("icon2/pin1.png", size=(24, 24), color="#000000")

                 # Default to Active initially (since defaults to is_pinned=True)
                 # toggle_pin will handle updates
                 self.btn_pin = tk.Label(self.header, image=self.icon_pin_active, bg="#444444", cursor="hand2")
             except Exception as e:
                 print("Error loading Pin icon: {}".format(e))
                 # Fallback to canvas if fails
                 self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg="#444444", highlightthickness=0)
                 self.draw_pin_icon()
             except Exception as e:
                 print("Error loading Pin icon: {}".format(e))
                 # Fallback to canvas if fails
                 self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg="#444444", highlightthickness=0)
                 self.draw_pin_icon() 
        else:
             self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg="#444444", highlightthickness=0)
             self.draw_pin_icon()
             
        self.btn_pin.pack(side="right", padx=5, pady=5)
        self.btn_pin.bind("<Button-1>", lambda e: self.toggle_pin())
        self.pin_tooltip = ToolTip(self.btn_pin, "Pin Window Open")
        
        # Custom Settings Button (Cog)
        if os.path.exists("icon2/spanner.png"):
            # Settings: 22x22 (Color: White)
            try:
                img = self.load_icon_colored("icon2/spanner.png", size=(22, 22), color="#FFFFFF")
                self.image_cache["settings_header"] = img
                self.btn_settings = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
        

            except Exception as e:
                 print("Error loading Spanner icon: {}".format(e))
                 self.btn_settings = tk.Label(self.header, text="⚙", bg="#444444", fg="#aaaaaa", font=(self.font_family, 12), cursor="hand2")
        else:
            self.btn_settings = tk.Label(self.header, text="⚙", bg="#444444", fg="#aaaaaa", font=(self.font_family, 12), cursor="hand2")
        self.btn_settings.pack(side="right", padx=5)
        self.btn_settings.bind("<Button-1>", lambda e: self.open_settings())

        # Help Button (?)
        self.btn_help = tk.Label(self.header, text="?", bg="#444444", fg="#AAAAAA", font=("Segoe UI", 14, "bold"), cursor="hand2")
        self.btn_help.pack(side="right", padx=(5, 5), pady=5)
        self.btn_help.bind("<Button-1>", lambda e: self.toggle_help_panel())
        ToolTip(self.btn_help, "Instructions")

        # Refresh Button
        if os.path.exists("icon2/refresh.png"):
            # Refresh: 22x22 (Color: White)
            try:
                img = self.load_icon_colored("icon2/refresh.png", size=(22, 22), color="#FFFFFF")
                self.image_cache["sync_header"] = img
                self.btn_refresh = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
            except Exception as e:
                 print("Error loading Refresh icon: {}".format(e))
                 self.btn_refresh = tk.Label(self.header, text="↻", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        else:
            self.btn_refresh = tk.Label(self.header, text="↻", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        self.btn_refresh.pack(side="right", padx=5)
        self.btn_refresh.bind("<Button-1>", lambda e: self.refresh_emails())
        
        ToolTip(self.btn_settings, "Settings")
        ToolTip(self.btn_refresh, "Refresh Email List")

        # Share Button
        if os.path.exists("icon2/share.png"):
            # Share: 20x20 (Color: White)
            try:
                img = self.load_icon_colored("icon2/share.png", size=(20, 20), color="#FFFFFF")
                self.image_cache["share_header"] = img
                self.btn_share = tk.Label(self.header, image=img, bg="#444444", cursor="hand2")
            except Exception as e:
                 print("Error loading Share icon: {}".format(e))
                 self.btn_share = tk.Label(self.header, text="🔗", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        else:
            self.btn_share = tk.Label(self.header, text="🔗", bg="#444444", fg="#aaaaaa", font=(self.font_family, 15), cursor="hand2")
        self.btn_share.pack(side="right", padx=5)
        # No action yet, just tooltip
        ToolTip(self.btn_share, "Sharing not available yet")

        # Content Area - Using PanedWindow for draggable resizing
        # Replaces grid_container
        
        # PanedWindow
        # sashwidth=4, sashrelief="raised" or "flat", bg="#333333" for visibility
        # opaqueresize=False preventing constant redraw during drag (smoother)
        self.paned_window = tk.PanedWindow(self.main_frame, orient="vertical", bg="#333333", sashwidth=4, sashrelief="flat", opaqueresize=False)
        self.paned_window.pack(expand=True, fill="both", padx=5, pady=5)
        
        # Pane 1: Email Container
        self.pane_emails = tk.Frame(self.paned_window, bg="#222222")
        self.paned_window.add(self.pane_emails, minsize=100)
        
        # Pane 2: Reminder Container
        self.pane_reminders = tk.Frame(self.paned_window, bg="#222222")
        self.paned_window.add(self.pane_reminders, minsize=100)
        
        # Note: We need to set the sash position AFTER the window is rendered/updated.
        # We'll do this in refresh_emails or a delayed call.
        
        # Email section header (Created once, kept in pane_emails)
        email_header = tk.Frame(self.pane_emails, bg="#333333", height=26)
        email_header.pack(fill="x", side="top")
        email_header.pack_propagate(False)  # Maintain fixed height
        
        # Email List Container (Scrollable) - Will be filled in refresh_emails
        self.email_list_frame = tk.Frame(self.pane_emails, bg="#222222")
        self.email_list_frame.pack(fill="both", expand=True)

        # ------------------
        # Header Controls
        # ------------------
        
        self.lbl_email_header = tk.Label(
            email_header, text="Email", 
            bg="#333333", fg="#FFFFFF",
            font=("Segoe UI", 9, "bold"), anchor="w"
        )
        self.lbl_email_header.pack(side="left", padx=5)
        
        # Dropdown / Drawer Button (Arrow)
        # Load arrow icon (Down = Closed, Up = Open)
        # Note: User requested "arrow24 icon (rotated 180 degrees)" for Closed state?
        # Standard: Down arrow usually means "Expand/Open" or "Below". 
        # User said: "rotated 180 degrees on the RHS... when open the arrow can rotate so it's facing up"
        # If the file 'arrow 24.png' is a DOWN arrow, then 180 is UP.
        # If 'arrow 24.png' is RIGHT arrow, 180 is LEFT.
        # Let's assume standard resource is 'arrow 24.png' (implied Right or Down?). 
        # Let's verify by loading. 
        # Actually I will just load and rotate dynamically.
        
        # Dropdown / Drawer Button (Arrow)
        # Dynamic Generation for "Infilled White" arrow as requested
        try:
             # Size 20x20
             # DOWN Arrow (Closed)
             img_down = Image.new("RGBA", (20, 20), (0,0,0,0))
             draw_d = ImageDraw.Draw(img_down)
             # Triangle pointing down: Top-Left, Top-Right, Bottom-Center
             # Coordinates for 20x20: (4, 7), (16, 7), (10, 15)
             draw_d.polygon([(4, 7), (16, 7), (10, 15)], fill="white")
             self.icon_arrow_down = ImageTk.PhotoImage(img_down)
             
             # UP Arrow (Open)
             img_up = Image.new("RGBA", (20, 20), (0,0,0,0))
             draw_u = ImageDraw.Draw(img_up)
             # Triangle pointing up: Bottom-Left, Bottom-Right, Top-Center
             # Coordinates: (4, 15), (16, 15), (10, 7)
             draw_u.polygon([(4, 15), (16, 15), (10, 7)], fill="white")
             self.icon_arrow_up = ImageTk.PhotoImage(img_up)
             
             self.btn_account_toggle = tk.Label(email_header, image=self.icon_arrow_down, bg="#333333", cursor="hand2")
        except Exception as e:
             print("Error generating arrow icons: {}".format(e))
             self.btn_account_toggle = tk.Label(email_header, text="▼", bg="#333333", fg="white", cursor="hand2")
             
        self.btn_account_toggle.pack(side="right", padx=5)
        self.btn_account_toggle.bind("<Button-1>", lambda e: self.toggle_account_selection())
            

        
        # ------------------

        
        self.scroll_frame = ScrollableFrame(self.email_list_frame, bg="#222222")
        self.scroll_frame.pack(expand=True, fill="both")
        
        # Reminder List Setup (Initial empty state, populated in refresh_reminders)
        # Note: We already added pane_reminders to PanedWindow
        
        # Header inside pane_reminders (for consistency with email pane)
        r_header = tk.Frame(self.pane_reminders, bg="#333333", height=20)
        r_header.pack(fill="x", side="top")
        r_header.pack_propagate(False)
        
        tk.Label(r_header, text="Flagged/Reminders", 
                 bg="#333333", fg="#AAAAAA", font=(self.font_family, 9, "bold")
        ).pack(side="left", padx=10, pady=3)

        self.reminder_list = ScrollableFrame(self.pane_reminders, bg="#1e1e1e")
        self.reminder_list.pack(fill="both", expand=True)
        
        # Reminder placeholder removed
        pass

        # Resize Grip (Overlay on the right edge)
        self.resize_grip = tk.Frame(self.main_frame, bg="#666666", cursor="sb_h_double_arrow", width=5)
        self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
        self.resize_grip.bind("<B1-Motion>", self.on_resize_drag)
        self.resize_grip.bind("<ButtonRelease-1>", self.on_resize_release)

        # Hot Strip Visual overlay (only visible when collected)
        # We use a Canvas now to draw the animation
        self.hot_strip_canvas = tk.Canvas(self.main_frame, bg="#444444", highlightthickness=0)
        
        # --- Events ---
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Motion>", self.on_motion) 

        # Initial Load
        self.refresh_emails()
        
        # Initial State
        self.apply_state()
        self.apply_window_layout()  # Apply window mode (single/dual)
        
        # Start Background Polling
        self.start_polling()

    def quit_application(self):
        """Terminates the application."""
        self.destroy()
        sys.exit(0)



    def open_settings(self):
        """Toggle the settings panel."""
        self.toggle_settings_panel()

    def toggle_settings_panel(self):
        """Show or hide the settings panel alongside the email list."""
        # Ensure Help panel is closed
        if self.help_panel_open:
             self.toggle_help_panel()

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
            
            try:
                # Open the panel alongside email list
                self.settings_panel = SettingsPanel(self.content_wrapper, self, self.refresh_emails)
                self.settings_panel.pack(side="left", fill="y")
                self.settings_panel_open = True
                
                # Expand window by exactly +370px (Updated width)
                new_width = self.expanded_width + self.settings_panel.panel_width
                self.set_geometry(new_width)
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("Settings Error", "Failed to open settings:\n{}".format(e))
                
                # Revert UI changes
                self.main_frame.pack_propagate(True)
                self.main_frame.config(width=0)
                self.main_frame.pack(side="left", fill="both", expand=True)
                self.settings_panel_open = False

    def toggle_help_panel(self):
        """Show or hide the help panel alongside the email list."""
        # Ensure Settings panel is closed
        if self.settings_panel_open:
             self.toggle_settings_panel()

        if self.help_panel_open:
            # Close the panel
            if self.help_panel:
                self.help_panel.pack_forget()
                self.help_panel.destroy()
                self.help_panel = None
            self.help_panel_open = False
            
            # Unfreeze main_frame
            self.main_frame.pack_propagate(True)
            self.main_frame.config(width=0)
            self.main_frame.pack(side="left", fill="both", expand=True)
            
            # Restore original width
            self.set_geometry(self.expanded_width)
        else:
            # Freeze main_frame
            current_width = self.main_frame.winfo_width()
            self.main_frame.config(width=current_width)
            self.main_frame.pack_propagate(False)
            self.main_frame.pack(side="left", fill="y", expand=False)
            
            # Open the panel
            self.help_panel = HelpPanel(self.content_wrapper, self)
            self.help_panel.pack(side="left", fill="y")
            self.help_panel_open = True
            
            # Expand window
            new_width = self.expanded_width + self.help_panel.panel_width
            self.set_geometry(new_width)
        
    def load_icon_colored(self, path, size=None, color="#BFBFBF", is_rgb_tuple=False):
        """Loads an image, applies a solid color mask, and returns ImageTk.PhotoImage."""
        try:
            pil_img = Image.open(path).convert("RGBA")
            
            # Resize if needed
            if size:
                pil_img = pil_img.resize(size, RESAMPLE_MODE)
            
            # Determine color tuple (R, G, B)
            if is_rgb_tuple:
                target_color = color # Expected tuple (r,g,b)
            else:
                # Parse Hex string
                c = color.lstrip('#')
                target_color = tuple(int(c[i:i+2], 16) for i in (0, 2, 4))

            # Create solid color image
            colored_img = Image.new("RGBA", pil_img.size, target_color + (255,))
            
            # Extract Alpha channel from source
            r, g, b, a = pil_img.split()
            
            # Simple threshold mask logic from before, but applied to the colored block
            # mask = a.point(lambda p: 255 if p > 20 else 0) 
            
            # Use original alpha as mask for smoother edges (antialiasing), 
            # OR use valid alpha only.
            
            final_img = Image.new("RGBA", pil_img.size, (0, 0, 0, 0))
            final_img.paste(colored_img, (0, 0), mask=a)
            
            return ImageTk.PhotoImage(final_img)
        except Exception as e:
            print("Error loading/coloring icon {}: {}".format(path, e))
            return None

    def load_icon_white(self, path, size=None):
        """Legacy wrapper for load_icon_colored (defaults to standard grey)."""
        return self.load_icon_colored(path, size, color="#BFBFBF")

    def handle_custom_action(self, config, email_data):
        """Executes the selected actions on the specific email."""
        print("Executing Actions for {} on {}".format(config.get('label'), email_data.get('subject')))
        
        entry_id = email_data.get("entry_id")
        store_id = email_data.get("store_id") # Support multi-account
        if not entry_id:
            print("No EntryID found.")
            return

        item = self.outlook_client.get_item_by_entryid(entry_id, store_id)
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
                        else: print("Folder '{}' not found.".format(folder_name))
            except Exception as e:
                print("Error executing {}: {}".format(act_name, e))

        try:
            # Execute Action 1
            execute_single_action(config.get("action1"), config.get("folder"))
            
            # Execute Action 2 - REMOVED
            # execute_single_action(config.get("action2"), config.get("folder"))
                
            # Refresh UI
            self.after(500, self.refresh_emails)
            
        except Exception as e:
            print("Action execution loop error: {}".format(e))

    def toggle_card_actions(self, action_frame):
        if action_frame.winfo_viewable():
            action_frame.pack_forget()
        else:
            action_frame.pack(fill="x", pady=(5, 0))


        




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
        Uses multiple methods (SwitchToThisWindow, AttachThreadInput, SetWindowPos) to force focus.
        """
        if not hwnd:
            return False
        
        try:
            # 1. Force Restore if Minimized
            if user32.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, 9) # SW_RESTORE
            else:
                win32gui.ShowWindow(hwnd, 5) # SW_SHOW
            
            # 2. Try SwitchToThisWindow (Undocumented but powerful - Alt-Tab equivalent)
            try:
                user32.SwitchToThisWindow(hwnd, True)
            except:
                pass

            # 3. AttachThreadInput Magic (The "Strong" Method)
            # This allows us to inject into the thread queue of the target window to force foreground
            try:
                target_tid = user32.GetWindowThreadProcessId(hwnd, None)
                current_tid = kernel32.GetCurrentThreadId()
                
                if target_tid != current_tid:
                    # Attach
                    user32.AttachThreadInput(current_tid, target_tid, True)
                    try:
                        # Force Foreground
                        win32gui.SetForegroundWindow(hwnd)
                        
                        # Z-Order Trick: Top, then Not Top
                        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 0x0003) # SWP_NOMOVE | SWP_NOSIZE
                        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, 0x0003)
                    finally:
                        # Always Detach
                        user32.AttachThreadInput(current_tid, target_tid, False)
                else:
                    # Same thread
                    win32gui.SetForegroundWindow(hwnd)
                    
            except Exception as e:
                print("Focus Magic Error: {}".format(e))
                # Fallback to standard
                try: win32gui.SetForegroundWindow(hwnd)
                except: pass

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
                self.show_has_attachment = data.get("show_has_attachment", False)
                self.only_flagged = data.get("only_flagged", False)
                self.include_read_flagged = data.get("include_read_flagged", True)
                self.flag_date_filter = data.get("flag_date_filter", "Anytime")
                self.window_mode = data.get("window_mode", "single")
                self.enabled_accounts = data.get("enabled_accounts", {})
                # Reminder filters
                self.reminder_show_flagged = data.get("reminder_show_flagged", True)
                self.reminder_due_filter = data.get("reminder_due_filter", "Anytime")
                self.reminder_show_categorized = data.get("reminder_show_categorized", False)
                self.reminder_categories = data.get("reminder_categories", [])
                self.reminder_high_importance = data.get("reminder_high_importance", False)
                self.reminder_normal_importance = data.get("reminder_normal_importance", False)
                self.reminder_low_importance = data.get("reminder_low_importance", False)
                self.reminder_pending_meetings = data.get("reminder_pending_meetings", False)
                self.reminder_accepted_meetings = data.get("reminder_accepted_meetings", True)
                self.reminder_declined_meetings = data.get("reminder_declined_meetings", False)
                self.reminder_tasks = data.get("reminder_tasks", False)
                self.reminder_todo = data.get("reminder_todo", False)
                self.reminder_has_reminder = data.get("reminder_has_reminder", False)
                self.buttons_on_hover = data.get("buttons_on_hover", False)
                self.email_double_click = data.get("email_double_click", False)
        except (FileNotFoundError, ValueError, IndexError):
            # If config is missing or corrupt, auto-enable all discovered accounts
            try:
                available = self.outlook_client.get_accounts()
                if available:
                    for acc in available:
                        self.enabled_accounts[acc] = {"email": True, "calendar": True, "tasks": True}
                    self.save_config()
            except:
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
            "show_has_attachment": self.show_has_attachment,
            "only_flagged": self.only_flagged,
            "include_read_flagged": self.include_read_flagged,
            "flag_date_filter": self.flag_date_filter,
            "window_mode": self.window_mode,
            # Reminder filters
            "reminder_show_flagged": self.reminder_show_flagged,
            "reminder_due_filter": self.reminder_due_filter,
            "reminder_show_categorized": self.reminder_show_categorized,
            "reminder_categories": self.reminder_categories,
            "reminder_high_importance": self.reminder_high_importance,
            "reminder_normal_importance": self.reminder_normal_importance,
            "reminder_low_importance": self.reminder_low_importance,
            "reminder_pending_meetings": self.reminder_pending_meetings,
            "reminder_accepted_meetings": self.reminder_accepted_meetings,
            "reminder_declined_meetings": self.reminder_declined_meetings,
            "reminder_tasks": self.reminder_tasks,
            "reminder_todo": self.reminder_todo,
            "reminder_has_reminder": self.reminder_has_reminder,
            "buttons_on_hover": self.buttons_on_hover,
            "email_double_click": self.email_double_click,
            "enabled_accounts": self.enabled_accounts
        }
        with open("sidebar_config.json", "w") as f:
            json.dump(data, f)

    def apply_window_layout(self):
        """Apply the current window mode (single or dual) to the layout."""
        if self.window_mode == "single":
            # Single window mode - hide reminder section
            try:
                self.paned_window.forget(self.pane_reminders)
            except:
                pass
        else:  # dual
            # Dual window mode - show reminder section
            # Check if likely already added
            if self.pane_reminders not in self.paned_window.panes():
                self.paned_window.add(self.pane_reminders, minsize=100)
            
            self.refresh_reminders()

    def flash_widget_recursive(self, widget, flash_color="#FFFFFF", duration=200):
        """Flashes a widget and ALL its children recursively."""
        try:
            # print("DEBUG: Flashing {}".format(widget))
            restore_map = {}
            
            def collect_and_flash(w):
                try:
                    if w.winfo_class() in ("Frame", "Label", "Canvas", "Text", "Button"):
                        orig = w.cget("bg")
                        restore_map[w] = orig
                        w.config(bg=flash_color)
                except Exception as e:
                    # print("DEBUG: Flash config failed for {}: {}".format(w, e))
                    pass
                
                # Recurse
                for child in w.winfo_children():
                    collect_and_flash(child)
            
            collect_and_flash(widget)
            
            # Revert
            self.after(duration, lambda: self._revert_flash(restore_map))
            
        except Exception as e:
            print("Flash Error: {}".format(e))

    def _revert_flash(self, restore_map):
        for w, orig in restore_map.items():
            try:
                w.config(bg=orig)
            except:
                pass

    def open_email(self, entry_id, source_widget=None):
        """Opens the specific email item with visual feedback."""
        if source_widget:
            self.flash_widget_recursive(source_widget)
                
        try:
             if not self.outlook_client.namespace:
                 self.outlook_client.connect()
             
             if self.outlook_client.namespace:
                 item = self.outlook_client.namespace.GetItemFromID(entry_id)
                 item.Display()
                 
                 try:
                     inspector = item.GetInspector
                     inspector.Activate()
                     
                     # Force window usage if possible
                     try:
                        # Some versions of Outlook don't expose HWND on Inspector easily via OOM in Py2.7
                        # But we can try to find it via the caption immediately.
                        caption = inspector.Caption
                        # Start polling for it
                        self.after(50, lambda: self._wait_and_focus(caption, attempt=1))
                     except:
                        pass
                 except Exception as e:
                     print("Focus preparation error: {}".format(e))
             else:
                 print("Error: Not connected to Outlook")
        except Exception as e:
             print("Error opening email: {}".format(e))

    def toggle_body_hover(self, event, preview_label, show):
        """Shows or hides the email body preview on hover."""
        if not self.show_hover_content: return
        if self.email_show_body: return # Already permanent

        try:
             if show:
                 # Show: Pack it if not packed
                 if not preview_label.winfo_ismapped():
                      preview_label.pack(fill="x", padx=5, pady=(0, 2))
             else:
                 # Hide: Unpack
                 if preview_label.winfo_ismapped():
                      preview_label.pack_forget()
        except:
             pass

    def _wait_and_focus(self, title_fragment, attempt=1):
        """Polls for window and forces focus using AttachThreadInput."""
        if attempt > 15:
            # print("DEBUG: Could not find window with title '{}'".format(title_fragment))
            return

        found_hwnd = None
        
        def callback(hwnd, ctx):
            if not win32gui.IsWindowVisible(hwnd): return
            txt = win32gui.GetWindowText(hwnd)
            if title_fragment in txt:
                ctx.append(hwnd)

        wins = []
        try: win32gui.EnumWindows(callback, wins)
        except: pass
             
        if wins:
            target_hwnd = wins[0]
            # print("DEBUG: Found window {} for '{}'".format(hex(target_hwnd), title_fragment))
            self._focus_window_by_hwnd(target_hwnd)
        else:
            self.after(100, lambda: self._wait_and_focus(title_fragment, attempt+1))

    def toggle_account_selection(self):
        """Toggles the account selection overlay."""
        if hasattr(self, "account_overlay") and self.account_overlay and self.account_overlay.winfo_exists():
            # Closing: Save and Destroy
            if hasattr(self, "account_ui_helper"):
                new_settings = self.account_ui_helper.get_settings()
                self.enabled_accounts = new_settings
                self.save_config()
                self.refresh_emails()
                self.refresh_reminders()
                
            self.account_overlay.destroy()
            self.account_overlay = None
            
            # Rotate Icon Down (Closed)
            if hasattr(self, "icon_arrow_down"):
                self.btn_account_toggle.config(image=self.icon_arrow_down)
        else:
            # Opening
            # Rotate Icon Up (Open)
            if hasattr(self, "icon_arrow_up"):
                self.btn_account_toggle.config(image=self.icon_arrow_up)
            
            # Fetch Accounts
            accounts = self.outlook_client.get_accounts()
            if not accounts:
                 messagebox.showerror("Error", "Could not fetch Outlook accounts.")
                 return

            # Note: We want to overlay the 'email_list_frame'.
            # But 'place' is relative to the parent. 'email_list_frame' parent is 'pane_emails'.
            # We want to cover 'pane_emails' but maybe start below the header?
            # 'email_header' is 20px high.
            
            target_frame = self.main_frame
            
            self.account_overlay = tk.Frame(target_frame, bg="#202020")
            
            # Geometry: Cover entire main frame (including footer)
            # Adjust y if needed to avoid covering custom title bar if it's inside main_frame
            # Assuming main_frame is the body. The Header is likely outside or at top.
            # We'll start at y=0 relative to main_frame.
            # Geometry: Cover main frame, but start BELOW the email header
            # Main Header (40) + Pad (5) + Email Header (26) + 1 = 72
            # This shows "Mail Mate" AND "Email - X" bars.
            self.account_overlay.place(x=0, y=72, relwidth=1.0, relheight=1.0, height=-72)
            
            # Add UI
            # We don't need the footer with buttons, so just use AccountSelectionUI
            # We use 'self.launch_folder_selection_from_overlay' as callback
            self.account_ui_helper = AccountSelectionUI(
                self.account_overlay, 
                accounts, 
                self.enabled_accounts, 
                self.launch_folder_selection_from_overlay, 
                bg_color="#202020"
            )
            self.account_ui_helper.pack(fill="both", expand=True)
            
            # Raise to top
            self.account_overlay.lift()

    def launch_folder_selection_from_overlay(self, account_name, on_selected, selected_paths=None):
         """Callback for folder picker spawned from overlay."""
         folders = self.outlook_client.get_folder_list(account_name)
         if not folders:
             messagebox.showwarning("No Folders", "Could not retrieve folder list for '{}'.".format(account_name))
             return
         
         def on_return():
             if hasattr(self, "overlay_picker"):
                 self.overlay_picker.destroy()
             if hasattr(self, "account_ui_helper"):
                 self.account_ui_helper.pack(fill="both", expand=True)

         def on_pick(paths):
             on_selected(paths)
             
         # Switch View
         if hasattr(self, "account_ui_helper"):
             self.account_ui_helper.pack_forget()
             
         container = self.account_overlay
         self.overlay_picker = FolderPickerFrame(container, folders, on_pick, on_return, selected_paths)
         self.overlay_picker.pack(fill="both", expand=True)

    def refresh_emails(self):
        try:
            # Update UI fonts for header elements
            self.lbl_title.config(font=(self.font_family, 10, "bold"))
            self.btn_settings.config(font=(self.font_family, 12))
            self.btn_refresh.config(font=(self.font_family, 15))

            # Clear existing
            for widget in self.scroll_frame.scrollable_frame.winfo_children():
                widget.destroy()

            # Determine enabled accounts
            accounts = [n for n, s in self.enabled_accounts.items() if s.get("email")] if self.enabled_accounts else None

            emails, unread_count = self.outlook_client.get_inbox_items(
                count=30, 
                unread_only=not self.show_read,
                account_names=accounts,
                account_config=self.enabled_accounts
            )
            
            # Update Header Count
            try:
                 self.lbl_email_header.config(text="Email - {}".format(unread_count))
            except: pass
            
            # Fetch Category Colors
            cat_map = self.outlook_client.get_category_map()
            
            for email in emails:
                lbl_sender = None
                lbl_subject = None
                lbl_preview = None
                
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
                                badge_text = "FLAGGED {}D".format(diff)
                            badge_bg = "#8E8E8E"
                        except:
                            pass

                header_frame = tk.Frame(card, bg=bg_color)
                header_frame.pack(fill="x")

                # Sender
                if self.email_show_sender:
                    sender_text = email['sender']
                    if is_unread:
                        sender_text = u"● " + sender_text # Add indicator dot
                        
                    lbl_sender = tk.Label(
                        header_frame, 
                        text=sender_text, 
                        fg="white", 
                        bg=bg_color, 
                        font=(self.font_family, self.font_size, "bold"),
                        anchor="w"
                    )
                    lbl_sender.pack(side="left", fill="x", expand=True)


                # Attachment indicator (only show if setting is enabled)
                if email.get('has_attachments', False) and self.show_has_attachment:
                    lbl_attachment = tk.Label(
                        header_frame, 
                        text="@", 
                        fg="#60CDFF", 
                        bg=bg_color, 
                        font=(self.font_family, self.font_size + 1, "bold"),
                    )
                    lbl_attachment.pack(side="right", padx=(4, 2))

                # Importance Indicator (High/Low)
                importance_val = email.get('importance', 1) # 0=Low, 1=Normal, 2=High
                if importance_val != 1:
                    imp_text = "!"
                    # High = Red-ish, Low = Grey
                    imp_fg = "#FF5555" if importance_val == 2 else "#AAAAAA" 
                    
                    lbl_importance = tk.Label(
                        header_frame, 
                        text=imp_text, 
                        fg=imp_fg, 
                        bg=bg_color, 
                        font=(self.font_family, self.font_size + 1, "bold"),
                    )
                    lbl_importance.pack(side="right", padx=(0, 2))

                # Categories Indicators
                categories_str = email.get('categories', "")
                if categories_str:
                    # Split and show badges
                    # Categories can be comma or semicolon separated
                    cats = re.split(r'[;,]', categories_str)
                    for cat in cats:
                        cat = cat.strip()
                        if not cat: continue
                    for cat in cats:
                        cat = cat.strip()
                        if not cat: continue
                        
                        # Lookup color
                        badge_bg = cat_map.get(cat, "#444444")
                        if badge_bg in ["#FFF768", "#F0E16C", "#EAC389"]: # Light colors
                            badge_fg = "#222222"
                        else:
                            badge_fg = "#FFFFFF"

                        # Just the color block
                        lbl_cat = tk.Frame(
                            header_frame, 
                            bg=badge_bg, 
                            width=10,
                            height=10
                        )
                        lbl_cat.pack(side="right", padx=1, pady=2)
                        
                        # Tooltip for the name
                        ToolTip(lbl_cat, cat)

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
                if self.email_show_subject:
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
                # Create if either Permanent Show OR Hover Show is enabled
                lbl_preview = None
                # Capture current body lines setting for this card
                try: 
                    lines = int(self.email_body_lines)
                except: 
                    lines = 2
                    
                if self.email_show_body or self.show_hover_content:
                    lbl_preview = tk.Text(
                        card, 
                        height=lines,
                        bg=bg_color, 
                        fg="#999999", 
                        font=(self.font_family, self.font_size - 1),
                        bd=0,
                        highlightthickness=0,
                        wrap="word",
                        cursor="arrow"
                    )
                    # Get preview text or fallback
                    preview_text = email.get('body', '').strip() 
                    if not preview_text:
                        preview_text = "(No preview available)"
                    
                    lbl_preview.insert("1.0", preview_text)
                    lbl_preview.config(state="disabled") # Read-only
                    
                    # Check if we should initially pack it (Show Body = True)
                    if self.email_show_body:
                         lbl_preview.pack(fill="x")
                
                # --- Action Frame (Buttons) ---
                # Rename locally to frame_buttons to match references
                frame_buttons = tk.Frame(card, bg=bg_color)
                
                # Populate buttons first (so they exist for binding)
                # Filter for valid buttons (Must have Icon AND Action)
                valid_buttons = [
                    conf for conf in self.btn_config 
                    if conf.get("icon") and conf.get("action1") != "None"
                ]

                for conf in valid_buttons:
                    icon = conf.get("icon", "🔘")
                    
                    is_png = icon.lower().endswith(".png")
                    btn_image = None
                    
                    if is_png:
                        if icon in self.image_cache:
                            btn_image = self.image_cache[icon]
                        else:
                            path = os.path.join("icons", icon)
                            if os.path.exists(path):
                                btn_image = self.load_icon_white(path, size=(24, 24))
                                if btn_image:
                                    self.image_cache[icon] = btn_image
                    
                    if btn_image:
                        btn = tk.Label(
                            frame_buttons, 
                            image=btn_image, 
                            bg=bg_color,
                            padx=10, pady=5,
                            cursor="hand2"
                        )
                    else:
                        btn = tk.Label(
                            frame_buttons, 
                            text=icon, 
                            fg="white", 
                            bg=bg_color,
                            font=("Segoe UI", 12),
                            padx=10, pady=5,
                            cursor="hand2"
                        )
                    
                    if len(valid_buttons) == 1:
                        btn.pack(side="left", expand=True, fill="y", ipadx=20)
                    else:
                        btn.pack(side="left", expand=True, fill="both")
                    
                    # Button Styling Bindings
                    btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#444444"))
                    btn.bind("<Leave>", lambda e, b=btn, bg=bg_color: b.config(bg=bg))
                    
                    # Tooltip logic
                    act1 = conf.get("action1", "")
                    act2 = conf.get("action2", "None")
                    tip_text = "{} & {}".format(act1, act2) if act2 != "None" else act1
                    ToolTip(btn, tip_text)
                    
                    # Bind Action
                    btn.bind("<Button-1>", lambda e, c=conf, em=email: self.handle_custom_action(c, em))

                # --- Logic for Buttons Visibility ---
                if self.buttons_on_hover:
                    # Start hidden
                    frame_buttons.pack_forget()
                else:
                    # Always show
                    frame_buttons.pack(fill="x", expand=True, padx=2, pady=(0, 2))


                # --- HOVER BINDINGS (Content & Buttons) ---
                # Define common show/hide helpers with DEFAULT ARGS to capture loop variables correctly
                # We also capture 'lines' from the scope to ensure correct height
                def show_hover_elements(e, lp=lbl_preview, fb=frame_buttons, h=lines):
                    # 1. Show Body Preview if enabled and not permanent
                    if self.show_hover_content and not self.email_show_body and lp:
                         if not lp.winfo_ismapped():
                             lp.config(height=h) 
                             lp.pack(fill="x", padx=5, pady=(0, 2)) 
                    
                    # 2. Show Buttons if enabled
                    if self.buttons_on_hover:
                         if not fb.winfo_ismapped():
                              fb.pack(fill="x", expand=True, padx=2, pady=(0, 2))
                
                def hide_hover_elements(e, lp=lbl_preview, fb=frame_buttons):
                    # 1. Hide Body Preview
                    if self.show_hover_content and not self.email_show_body and lp:
                         if lp.winfo_ismapped():
                              lp.pack_forget()
                    
                    # 2. Hide Buttons
                    if self.buttons_on_hover:
                         if fb.winfo_ismapped():
                              fb.pack_forget()

                
                # Robust Hide Logic using winfo_containing
                def robust_hide(e, c=card, lp=lbl_preview, fb=frame_buttons):
                    # Cancel pending show
                    if hasattr(c, "_show_timer") and c._show_timer:
                        c.after_cancel(c._show_timer)
                        c._show_timer = None
                    
                    try:
                        x, y = c.winfo_pointerxy()
                        widget = c.winfo_containing(x, y)
                        # Stay shown if mouse is over card or any of its descendants
                        if not widget or (widget != c and not str(widget).startswith(str(c))):
                            hide_hover_elements(e, lp, fb)
                    except:
                        pass # Safety
                
                def safe_show(e, c=card, lp=lbl_preview, fb=frame_buttons):
                     # Delay show to prevent flashing (Debounce)
                     if hasattr(c, "_show_timer") and c._show_timer:
                         c.after_cancel(c._show_timer)
                     c._show_timer = c.after(250, lambda: show_hover_elements(e, lp, fb))

                # Apply Bindings
                if (self.show_hover_content and not self.email_show_body) or self.buttons_on_hover:
                    card.bind("<Enter>", safe_show)
                    card.bind("<Leave>", robust_hide)
                    
                    # Bind children to prevent flickering
                    for child in card.winfo_children():
                        child.bind("<Enter>", safe_show)
                        child.bind("<Leave>", robust_hide)
                
                # Standard Click (Open Email) logic for card
                # --- CLICK LOGIC (Open Email) ---
                def on_card_click(e, eid=email['entry_id'], w=card):
                    self.open_email(eid, source_widget=w)

                # Apply Click Bindings
                if self.email_double_click: 
                     card.bind("<Double-Button-1>", on_card_click)
                     card.bind("<Button-1>", lambda e, c=card: c.focus_set())
                else:
                     card.bind("<Button-1>", on_card_click)
                
                # Bind Children (Robustly)
                for child in card.winfo_children():
                     # Don't bind click to buttons (they have their own actions)
                     if child != frame_buttons and getattr(child, "master", None) != frame_buttons:
                        if self.email_double_click: 
                            child.bind("<Double-Button-1>", on_card_click)
                        else:
                            child.bind("<Button-1>", on_card_click)
                     # Preview text click -> Open Email
                     if child == lbl_preview:
                          if self.email_double_click: 
                              child.bind("<Double-Button-1>", on_card_click)
                          else:
                              child.bind("<Button-1>", on_card_click)

                
                if self.email_double_click: 
                      # Logic handled by bind_click helper inside loop (Wait, bind_click isn't shown here)
                      # Assuming bind_click handles double click check or we need to add it.
                      # The loop continues...
                     card.bind("<Double-Button-1>", on_card_click)
                     if lbl_sender: lbl_sender.bind("<Double-Button-1>", on_card_click)
                     if lbl_subject: lbl_subject.bind("<Double-Button-1>", on_card_click)
                     if lbl_preview: lbl_preview.bind("<Double-Button-1>", on_card_click)
                     
                     # Optional: Single click handles focus or selection
                     card.bind("<Button-1>", lambda e: card.focus_set())
                else:
                     # Standard Single Click
                     card.bind("<Button-1>", on_card_click)
                     if lbl_sender: lbl_sender.bind("<Button-1>", on_card_click)
                     if lbl_subject: lbl_subject.bind("<Button-1>", on_card_click)
                     if lbl_preview: lbl_preview.bind("<Button-1>", on_card_click) 
                
                # Dynamic wrapping for both labels
                def update_wraps(e, s=lbl_subject, p=lbl_preview):
                    width = e.width - 20
                    if s:
                        s.config(wraplength=width)
                    # Only wrap if it's a Label (Text widgets handle wrapping internally)
                    if p and isinstance(p, tk.Label):
                        p.config(wraplength=width)
                    
                card.bind("<Configure>", update_wraps)

            # Ensure Reminders are also refreshed
            self.refresh_reminders()

        except Exception as e:
            print("CRITICAL ERROR in refresh_emails: {}".format(e))
            import traceback
            traceback.print_exc()
            # Try to show error to user if possible
            try:
                messagebox.showerror("Sidebar Error", "Error refreshing emails:\\n{}".format(e))
            except: pass


    def refresh_reminders(self):
        """Refreshes the Reminder/Flagged section (Bottom List)."""
        # Ensure scrollable frame exists
        # Clear content
        if self.reminder_list:
            for widget in self.reminder_list.scrollable_frame.winfo_children():
                widget.destroy()
        
        container = self.reminder_list.scrollable_frame
        
        # Helper for binding click
        def bind_click(widget, entry_id):
            if self.email_double_click:
                widget.bind("<Double-Button-1>", lambda e, eid=entry_id, w=widget: self.open_email(eid, source_widget=w))
            else:
                widget.bind("<Button-1>", lambda e, eid=entry_id, w=widget: self.open_email(eid, source_widget=w))
        
        # 1. Meetings (Today & Tomorrow)
        # 1. Meetings
        now = datetime.now()
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        
        # Calculate End Date based on Selection
        # If multiple are selected, we take the MAX date
        end_date = today_start # Start with today
        
        has_date_filter = False
        
        if "Today" in self.reminder_meeting_dates:
             # Today EOD
             d = today_start + timedelta(days=1) - timedelta(seconds=1)
             if d > end_date: end_date = d
             has_date_filter = True
             
        if "Tomorrow" in self.reminder_meeting_dates:
             # Tomorrow EOD
             d = today_start + timedelta(days=2) - timedelta(seconds=1)
             if d > end_date: end_date = d
             has_date_filter = True

        if "Next 7 Days" in self.reminder_meeting_dates:
             d = today_start + timedelta(days=8) - timedelta(seconds=1) # Today + 7 full days
             if d > end_date: end_date = d
             has_date_filter = True
             
        if "Custom" in self.reminder_meeting_dates:
             try:
                 days = int(getattr(self, "reminder_custom_days", 30))
             except: days = 30
             d = today_start + timedelta(days=days+1) - timedelta(seconds=1)
             if d > end_date: end_date = d
             has_date_filter = True
             
        # If no date filter, maybe don't show any? Or default?
        # User said "defaults for next one should be Today, Tomorrow". 
        # If they untick all, implies show none?
        if not has_date_filter:
             # Show none
             meetings = []
        else:
             cal_accounts = [n for n, s in self.enabled_accounts.items() if s.get("calendar")] if self.enabled_accounts else None
             # Pass datetime objects directly
             raw_meetings = self.outlook_client.get_calendar_items(today_start, end_date, cal_accounts)
             
             # Filter by Status
             # olResponseNone = 0, olResponseOrganized = 1, olResponseTentative = 2, olResponseAccepted = 3, olResponseDeclined = 4
             meetings = []
             for m in raw_meetings:
                 status = m.get("response_status", 0)
                 
                 # Accepted
                 if status == 3 and self.reminder_accepted_meetings:
                     meetings.append(m)
                     continue
                     
                 # Declined
                 if status == 4 and self.reminder_declined_meetings:
                     meetings.append(m)
                     continue
                     
                 # Pending (None, Organized, Tentative, NotResponded=5)
                 # Basically anything not Accepted(3) or Declined(4)
                 if status not in [3, 4] and self.reminder_pending_meetings:
                     meetings.append(m)
                     continue
        
        if meetings:
            tk.Label(container, text="CALENDAR", fg="#60CDFF", bg="#1e1e1e", font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(5, 2))
            for m in meetings:
                 mf = tk.Frame(container, bg="#252526", padx=5, pady=5)
                 mf.pack(fill="x", padx=2, pady=1)
                 
                 # Time
                 # Time
                 try:
                     dt = m['start']
                     is_today = dt.date() == now.date()
                     if is_today:
                         time_str = dt.strftime("%I:%M %p")
                     else:
                         # Show "Tom 10:00 AM" or "Mon 10:00 AM"
                         is_tomorrow = dt.date() == (now.date() + timedelta(days=1))
                         if is_tomorrow:
                             time_str = "Tom " + dt.strftime("%I:%M %p")
                         else:
                             time_str = dt.strftime("%a %I:%M %p")
                 except:
                     time_str = "??"
                     
                 tk.Label(mf, text=time_str, fg="#AAAAAA", bg="#252526", font=("Segoe UI", 9)).pack(side="left")
                 subj = tk.Label(mf, text=m['subject'], fg="white", bg="#252526", font=("Segoe UI", 9, "bold"))
                 subj.pack(side="left", padx=5)
                 
                 bind_click(mf, m['entry_id'])
                 bind_click(subj, m['entry_id'])

                 # Calendar Buttons Frame (Hidden initially)
                 c_actions = tk.Frame(mf, bg="#252526")
                 # c_actions.pack(side="right", padx=2) # Hidden

                 # Helper to create buttons (Inline reuse or call similar logic)
                 def make_cal_btn(parent, text, cmd, tip):
                     btn = tk.Label(parent, text=text, fg="#AAAAAA", bg="#252526", font=("Segoe UI", 10), cursor="hand2", padx=5)
                     btn.pack(side="right", padx=5) # Align right
                     btn.bind("<Button-1>", lambda e: cmd())
                     btn.bind("<Enter>", lambda e: btn.config(fg="white", bg="#444444"))
                     btn.bind("<Leave>", lambda e: btn.config(fg="#AAAAAA", bg="#252526"))
                     if tip: ToolTip(btn, tip)
                     return btn

                 # Dismiss Button (Checkmark / Delete) - Far Right
                 def do_dismiss_cal(eid=m['entry_id'], w=mf):
                     # Confirm simple deletion/dismissal? User asked for "check box" logic.
                     # We'll just delete it for now as "Complete".
                     success = self.outlook_client.dismiss_calendar_item(eid)
                     if success:
                         w.pack_forget()
                     else:
                         messagebox.showerror("Error", "Failed to dismiss meeting.")
                         
                 # Try to load PNG
                 btn_dismiss = None
                     # Try to load PNG
                 btn_dismiss = None
                 if os.path.exists("icons/Delete.png"): # Use Delete icon for dismissal
                      img = self.load_icon_colored("icons/Delete.png", size=(24, 24), color="#AAAAAA") # Larger 24x24
                      if img:
                          btn_dismiss = tk.Label(c_actions, image=img, bg="#252526", cursor="hand2", padx=5)
                          btn_dismiss.image = img
                 
                 if not btn_dismiss:
                      btn_dismiss = make_cal_btn(c_actions, u"✓", do_dismiss_cal, "Dismiss/Delete")
                 else:
                      btn_dismiss.pack(side="right", padx=5)
                      btn_dismiss.bind("<Button-1>", lambda e: do_dismiss_cal())
                      if "Dismiss/Delete": ToolTip(btn_dismiss, "Dismiss/Delete")

                 # Open Button (Re-add using make_cal_btn or custom) - Left of Dismiss
                 make_cal_btn(c_actions, "📂", lambda eid=m['entry_id']: self.open_email(eid), "Open Meeting")
                 
                 # --- CALENDAR HOVER LOGIC ---
                 def show_c_actions(e, fa=c_actions):
                     if not fa.winfo_ismapped():
                         fa.pack(side="top", fill="x", padx=2, pady=(2, 0)) # Expand below
                 
                 def hide_c_actions(e, c=mf, fa=c_actions):
                     try:
                         x, y = c.winfo_pointerxy()
                         widget = c.winfo_containing(x, y)
                         if widget:
                             path = str(widget)
                             c_path = str(c)
                             if path.startswith(c_path): return
                     except: pass
                     
                     if fa.winfo_ismapped():
                         fa.pack_forget()

                 # Bindings
                 mf.bind("<Enter>", show_c_actions)
                 mf.bind("<Leave>", hide_c_actions)
                 subj.bind("<Enter>", show_c_actions)
                 subj.bind("<Leave>", hide_c_actions)

        # 2. Outlook Tasks
        if self.reminder_show_flagged:
             tasks = self.outlook_client.get_tasks(due_filters=self.reminder_due_filters, account_names=cal_accounts)
             
             if tasks:
                 tk.Label(container, text="TASKS", fg="#28a745", bg="#1e1e1e", font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(10, 2))
                 for task in tasks:
                     tf = tk.Frame(container, bg="#2d2d2d", highlightthickness=1, highlightbackground="#28a745", padx=5, pady=5)
                     tf.pack(fill="x", padx=2, pady=2)
                     
                     subj = tk.Label(tf, text=task['subject'], fg="white", bg="#2d2d2d", font=("Segoe UI", 9), anchor="w", justify="left", wraplength=self.expanded_width-40)
                     subj.pack(side="top", fill="x", expand=True, padx=5, pady=(0, 2))
                     
                     bind_click(tf, task['entry_id'])
                     bind_click(subj, task['entry_id'])

                     # Task Buttons Frame
                     t_actions = tk.Frame(tf, bg="#2d2d2d")
                     # t_actions.pack(side="right", padx=2) # Hidden

                     # Helper to create buttons
                     def make_task_btn(parent, text, cmd, tip):
                         btn = tk.Label(parent, text=text, fg="#AAAAAA", bg="#2d2d2d", font=("Segoe UI", 10), cursor="hand2", padx=5)
                         btn.pack(side="right", padx=5) # Align Right
                         btn.bind("<Button-1>", lambda e: cmd())
                         btn.bind("<Enter>", lambda e: btn.config(fg="white", bg="#444444"))
                         btn.bind("<Leave>", lambda e: btn.config(fg="#AAAAAA", bg="#2d2d2d"))
                         if tip: ToolTip(btn, tip)
                         return btn

                     # Complete Button (Checkmark) - Far Right
                     def do_complete(eid=task['entry_id'], w=tf):
                         success = self.outlook_client.mark_task_complete(eid)
                         if success:
                             # Fade out or remove
                             w.pack_forget()
                             # message?
                         else:
                             messagebox.showerror("Error", "Failed to mark task complete.")
                     
                     # Try PNG for complete
                     btn_complete = None
                     # Reuse 'Mark as Read' icon (often a check) or just unicode if PNG not suited
                     # 'icons/Mark as Read.png' exists.
                     if os.path.exists("icons/Mark as Read.png"):
                          img = self.load_icon_colored("icons/Mark as Read.png", size=(24, 24), color="#AAAAAA") # Large 24x24
                          if img:
                              btn_complete = tk.Label(t_actions, image=img, bg="#2d2d2d", cursor="hand2", padx=5)
                              btn_complete.image = img # Keep ref
                     
                     if not btn_complete:
                          make_task_btn(t_actions, u"✓", do_complete, "Mark Complete")
                     else:
                          btn_complete.pack(side="right", padx=5)
                          btn_complete.bind("<Button-1>", lambda e: do_complete())
                          ToolTip(btn_complete, "Mark Complete")

                     # Open Button (Folder icon or similar) - Left of Complete
                     btn_open = None
                     if os.path.exists("icon2/open-task.png"):
                          img = self.load_icon_colored("icon2/open-task.png", size=(24, 24), color="#AAAAAA") # Large 24x24
                          if img:
                              btn_open = tk.Label(t_actions, image=img, bg="#2d2d2d", cursor="hand2", padx=5)
                              btn_open.image = img
                     
                     if not btn_open:
                          make_task_btn(t_actions, u"📂", lambda eid=task['entry_id']: self.open_email(eid), "Open Task")
                     else:
                          btn_open.pack(side="right", padx=5)
                          btn_open.bind("<Button-1>", lambda e, eid=task['entry_id']: self.open_email(eid))
                          ToolTip(btn_open, "Open Task")
                     
                     # --- TASKS HOVER LOGIC ---
                     def show_t_actions(e, fa=t_actions):
                         if not fa.winfo_ismapped():
                             fa.pack(side="top", fill="x", padx=2, pady=(2, 0))
                     
                     def hide_t_actions(e, c=tf, fa=t_actions):
                         try:
                             x, y = c.winfo_pointerxy()
                             widget = c.winfo_containing(x, y)
                             if widget:
                                 path = str(widget)
                                 c_path = str(c)
                                 if path.startswith(c_path): return
                         except: pass
                         
                         if fa.winfo_ismapped():
                             fa.pack_forget()

                     tf.bind("<Enter>", show_t_actions)
                     tf.bind("<Leave>", hide_t_actions)
                     subj.bind("<Enter>", show_t_actions)
                     subj.bind("<Leave>", hide_t_actions)

        # 3. Flagged Emails
        if self.reminder_show_flagged:
             email_accounts = [n for n, s in self.enabled_accounts.items() if s.get("email")] if self.enabled_accounts else None
             flags, _ = self.outlook_client.get_inbox_items(
                 count=30,
                 unread_only=False,
                 only_flagged=True,
                 due_filters=self.reminder_due_filters,
                 account_names=email_accounts
             )
             
             if flags:
                 tk.Label(container, text="FLAGGED EMAILS", fg="#FF8C00", bg="#1e1e1e", font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(10, 2))
                 
                 for email in flags:
                     cf = tk.Frame(container, bg="#2d2d2d", highlightthickness=1, highlightbackground="#FF8C00", padx=5, pady=5)
                     cf.pack(fill="x", padx=2, pady=2)
                     
                     # Subject Label (Now packed top)
                     subj = tk.Label(cf, text=email['subject'], fg="white", bg="#2d2d2d", font=("Segoe UI", 9), anchor="w", justify="left", wraplength=self.expanded_width-40)
                     subj.pack(side="top", fill="x", expand=True, padx=5, pady=(0, 2))
                     
                     bind_click(cf, email['entry_id'])
                     bind_click(subj, email['entry_id'])

                     # Flag Actions Frame (Hidden initially)
                     # Packed below subject
                     f_actions = tk.Frame(cf, bg="#2d2d2d")
                     # f_actions.pack(side="top", fill="x", padx=2) # Hide by default

                     # Helper to create buttons
                     def make_flag_btn(parent, text, cmd, tip):
                         btn = tk.Label(parent, text=text, fg="#AAAAAA", bg="#2d2d2d", font=("Segoe UI", 10), cursor="hand2", padx=5)
                         btn.pack(side="right", padx=5) # Pack right for alignment
                         btn.bind("<Button-1>", lambda e: cmd())
                         btn.bind("<Enter>", lambda e: btn.config(fg="white", bg="#444444"))
                         btn.bind("<Leave>", lambda e: btn.config(fg="#AAAAAA", bg="#2d2d2d"))
                         if tip: ToolTip(btn, tip)
                         return btn

                     # Unflag Button (Flag icon) - Moved to far right (first packed right)
                     def do_unflag(eid=email['entry_id'], sid=email['store_id'], w=cf):
                         success = self.outlook_client.unflag_email(eid, sid)
                         if success:
                             w.pack_forget()
                         else:
                             messagebox.showerror("Error", "Failed to unflag email.")
                     
                     btn_unflag = None
                     if os.path.exists("icon2/flag.png"):
                          img = self.load_icon_colored("icon2/flag.png", size=(24, 24), color="#FF8C00") # 50% larger (16->24)
                          if img:
                              btn_unflag = tk.Label(f_actions, image=img, bg="#2d2d2d", cursor="hand2", padx=5)
                              btn_unflag.image = img
                     
                     if not btn_unflag:
                          make_flag_btn(f_actions, u"⚑", do_unflag, "Unflag")
                     else:
                          # Pack Right
                          btn_unflag.pack(side="right", padx=5)
                          btn_unflag.bind("<Button-1>", lambda e: do_unflag())
                          ToolTip(btn_unflag, "Unflag")

                     # Open Button (Folder icon)
                     btn_open = None
                     if os.path.exists("icon2/open-email.png"):
                          img = self.load_icon_colored("icon2/open-email.png", size=(24, 24), color="#AAAAAA") # 50% larger
                          if img:
                              btn_open = tk.Label(f_actions, image=img, bg="#2d2d2d", cursor="hand2", padx=5)
                              btn_open.image = img
                     
                     if not btn_open:
                          make_flag_btn(f_actions, u"📂", lambda eid=email['entry_id']: self.open_email(eid), "Open Email")
                     else:
                          # Pack Right (left/next to Unflag)
                          btn_open.pack(side="right", padx=5)
                          btn_open.bind("<Button-1>", lambda e, eid=email['entry_id']: self.open_email(eid))
                          ToolTip(btn_open, "Open Email")

                     # --- HOVER LOGIC ---
                     def show_f_actions(e, fa=f_actions):
                         if not fa.winfo_ismapped():
                             fa.pack(side="top", fill="x", padx=2, pady=(2, 0)) # Pack creates accordion expansion below subject
                     
                     def hide_f_actions(e, c=cf, fa=f_actions):
                         try:
                             x, y = c.winfo_pointerxy()
                             widget = c.winfo_containing(x, y)
                             # Check if we are still inside 'cf' or any of its children (fa included)
                             if widget:
                                 path = str(widget)
                                 c_path = str(c)
                                 if path.startswith(c_path): # widget is child of cf
                                     return
                         except: pass
                         
                         if fa.winfo_ismapped():
                             fa.pack_forget()

                     # Bind Enter/Leave on container and subject
                     cf.bind("<Enter>", show_f_actions)
                     cf.bind("<Leave>", hide_f_actions)
                     subj.bind("<Enter>", show_f_actions)
                     # No direct leave on subj needed if it propagates or handled by containing check
                     # But subj leave -> enters cf? Or enters void?
                     # Let's be safe:
                     subj.bind("<Leave>", hide_f_actions)
                     
                     # Also bind hover specifically for buttons area to prevent hiding?
                     # No, because buttons are children of fa, which is child of cf. containing check covers it.


    def draw_pin_icon(self):
        # Determine active icon based on state
        if isinstance(self.btn_pin, tk.Label):
             # If using Label with Images
             if self.is_pinned:
                 if hasattr(self, 'icon_pin_active'):
                     self.btn_pin.config(image=self.icon_pin_active)
             else:
                 if hasattr(self, 'icon_pin_inactive'):
                     self.btn_pin.config(image=self.icon_pin_inactive)
        
        elif isinstance(self.btn_pin, tk.Canvas):
            # If using Canvas drawing
            self.btn_pin.delete("all")
            color = "#007ACC" if self.is_pinned else "#AAAAAA"
            # Draw a simple pin shape
            self.btn_pin.create_oval(10, 5, 20, 15, fill=color, outline="")
            self.btn_pin.create_line(15, 15, 15, 25, fill=color, width=2)

    def toggle_pin(self):
        self.is_pinned = not self.is_pinned
        
        # Update Tooltip
        if self.is_pinned:
            if hasattr(self, 'icon_pin_active') and isinstance(self.btn_pin, tk.Label):
                 self.btn_pin.config(image=self.icon_pin_active)
            elif isinstance(self.btn_pin, tk.Canvas):
                 self.btn_pin.delete("all")
                 self.draw_pin_icon() # Redraw pinned state
            
            self.pin_tooltip.text = "Unpin Window (Current: Pinned)"
        else:
            if hasattr(self, 'icon_pin_inactive') and isinstance(self.btn_pin, tk.Label):
                 self.btn_pin.config(image=self.icon_pin_inactive)
            elif isinstance(self.btn_pin, tk.Canvas):
                 self.btn_pin.delete("all")
                 self.draw_pin_icon() # Redraw unpinned state (draw_pin_icon checks is_pinned)
            
            self.pin_tooltip.text = "Pin Window (Current: Auto-Collapse)"
        self.save_config()
        self.apply_state()
        
        # Force a check immediately to update pulse state
        self.after(100, self._perform_check)

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
            # Pinned: Reserve Full Width, Visual Full Width
            mode = "PINNED"
            reserve_w = self.expanded_width
            visual_w = self.expanded_width
            
        elif self.is_expanded:
            # Expanded/Overlay: Reserve Strip Width, Visual Full Width
            # This allows the "Overlay" effect without losing the "Sneak Behind" protection for the strip.
            mode = "OVERLAY"
            reserve_w = self.hot_strip_width
            visual_w = self.expanded_width
            
        else:
            # Collapsed: Reserve Strip Width, Visual Strip Width
            mode = "COLLAPSED"
            reserve_w = self.hot_strip_width
            visual_w = self.hot_strip_width

        # --- UI Management ---
        if mode == "COLLAPSED":
            # Hide internals
            self.header.pack_forget()
            self.resize_grip.place_forget()
            # Show Hot Strip
            self.hot_strip_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            # Show internals
            self.hot_strip_canvas.place_forget()
            self.header.pack(fill="x", side="top", before=self.paned_window)
            
            # Grip Placement
            if self.dock_side == "Left":
                self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
            else:
                self.resize_grip.place(relx=0.0, rely=0, anchor="nw", relheight=1.0)

        # --- AppBar & Geometry ---
        self.appbar.register() # Ensure always registered
        
        # 1. Set Reservation (This keeps other windows away from at least the strip)
        # Note: In Overlay mode, we only reserve the strip width, so we draw OVER other windows 
        # but they still respect the base strip edge.
        self.appbar.set_pos(reserve_w, self.monitor_x, self.monitor_y, self.screen_width, self.screen_height)
        
        # 2. Set Visual Geometry (Can be larger than reservation)
        self.set_geometry(visual_w)

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
        # Retrieve Monitor and Work Area info in one go
        # If pinned, we trust the AppBar check. If not, we recalc.
        
        # Current logic:
        # 1. Get Monitor Info based on current window position
        monitor_info = self.get_monitor_metrics()
        
        mx, my, mw, mh = monitor_info['monitor']
        wx, wy, ww, wh = monitor_info['work']
        
        # Determine Reference Geometry
        if self.appbar.registered:
            # Use the registered AppBar position (system adjusted)
            rect = self.appbar.abd.rc
            x = rect.left
            y = rect.top
            h = rect.bottom - rect.top
            
            if self.dock_side == "Right":
                 x = rect.right - width
        else:
            # Unpinned / Overlay Mode (Only if registration failed or deliberately unregistered)
            # Use Work Area for Height/Y to respect Taskbar
            y = wy
            h = wh
            
            # X Calculation: Anchor to Monitor Edge
            # But ensure we don't start 'under' a vertical taskbar on the left?
            # Using Work Area left (wx) is safest.
            
            if self.dock_side == "Left":
                x = wx
            else:
                x = wx + ww - width

        self.geometry("{}x{}+{}+{}".format(width, h, x, y))
        self.update_idletasks()
        self.wm_attributes("-topmost", True)

    def get_monitor_metrics(self):
        """
        Uses EnumDisplayMonitors to find the monitor closest to the window center.
        Returns check-safe dictionary with 'monitor' and 'work' tuples (x,y,w,h).
        """
    def get_monitor_metrics(self):
        """
        Uses EnumDisplayMonitors to find the monitor closest to the window center.
        Returns check-safe dictionary with 'monitor' and 'work' tuples (x,y,w,h).
        """
        if hasattr(self, 'hwnd') and self.hwnd:
            hwnd = self.hwnd
        else:
            hwnd = self.winfo_id()
            try:
                hwnd = ctypes.windll.user32.GetParent(hwnd) or hwnd
            except: pass
            
        # Get Window Rect center
        try:
             wr = wintypes.RECT()
             user32.GetWindowRect(hwnd, ctypes.byref(wr))
             cx = (wr.left + wr.right) // 2
             cy = (wr.top + wr.bottom) // 2
        except:
             # Fallback to screen center if window not visible
             cx = self.winfo_screenwidth() // 2
             cy = self.winfo_screenheight() // 2
             
        monitors = []

        def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            if user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi)):
                r = mi.rcMonitor
                w = mi.rcWork
                monitors.append({
                    'm_rect': (r.left, r.top, r.right, r.bottom),
                    'w_rect': (w.left, w.top, w.right, w.bottom)
                })
            return True

        MONITORENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.ULONG, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)
        user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0)
        
        best_mon = None
        min_dist = float('inf')
        
        for m in monitors:
            # Check if center is inside
            ml, mt, mr, mb = m['m_rect']
            if ml <= cx <= mr and mt <= cy <= mb:
                best_mon = m
                break
            
            # Distance to center
            # Simple Manhattan distance from monitor center to window center
            mcx = (ml + mr) // 2
            mcy = (mt + mb) // 2
            dist = abs(cx - mcx) + abs(cy - mcy)
            if dist < min_dist:
                min_dist = dist
                best_mon = m
                
        if not best_mon and monitors:
             best_mon = monitors[0] # Fallback to primary
             
        if best_mon:
             ml, mt, mr, mb = best_mon['m_rect']
             wl, wt, wr, wb = best_mon['w_rect']
             return {
                 'monitor': (ml, mt, mr - ml, mb - mt),
                 'work': (wl, wt, wr - wl, wb - wt)
             }
             
        # Ultimate fallback
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        return {
            'monitor': (0, 0, sw, sh),
            'work': (0, 0, sw, sh)
        }

    # Replaces get_current_monitor_info and get_work_area_info
    # Kept for compatibility if called elsewhere, but redirected
    def get_current_monitor_info(self):
        m = self.get_monitor_metrics()['monitor']
        return m
        
    def get_work_area_info(self):
        m = self.get_monitor_metrics()['work']
        return m

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
        self.geometry("+{}+{}".format(x, y))

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

    # --- Polling Control ---
    # --- Polling Control ---
    def start_polling(self):
        """Starts the background polling loop."""
        self.check_updates()
        self.check_fullscreen_app()

    def check_fullscreen_app(self):
        """Checks if a full-screen application is active on the current monitor and helps sidebar get out of the way."""
        try:
            # 1. Get Foreground Window
            hwnd_active = user32.GetForegroundWindow()
            if not hwnd_active:
                self.after(2000, self.check_fullscreen_app)
                return

            # 2. Get Class Name (Ignore Shell/Desktop)
            buff = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd_active, buff, 256)
            cls = buff.value
            
            on_same_monitor = False # Default
            
            # Ignore self and shell
            if cls in ["Progman", "WorkerW", "Shell_TrayWnd", "ImmersiveLauncher"] or hwnd_active == self.winfo_id():
                 # Not a "real" app we care about
                 # If we were suppressing, maybe we should restore? 
                 # Actually, if we are on Desktop, we probably want to Restore if we auto-collapsed.
                 pass
            else:
                # 3. Get Window Rect
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd_active, ctypes.byref(rect))
                fw = rect.right - rect.left
                fh = rect.bottom - rect.top
                
                # 4. Get Current Monitor Metrics for Sidebar
                metrics = self.get_monitor_metrics()
                mx, my, mw, mh = metrics['monitor']
                
                # 5. Check Intersection/Monitor
                # Does the active window center fall onto our monitor?
                fcx = (rect.left + rect.right) // 2
                fcy = (rect.top + rect.bottom) // 2
                
                on_same_monitor = (mx <= fcx <= mx + mw) and (my <= fcy <= my + mh)
                
                if on_same_monitor:
                    # 6. Check Dimensions (Allow small variance)
                    # Is it basically full monitor size?
                    is_fullscreen = (abs(fw - mw) < 20) and (abs(fh - mh) < 20)
                    
                    # print("DEBUG: Window '{}' Class: {} Size: {}x{} Monitor: {}x{} Fullscreen: {}".format(win32gui.GetWindowText(hwnd_active), cls, fw, fh, mw, mh, is_fullscreen))

                    if is_fullscreen:
                        # ACTIVE FULLSCREEN DETECTED
                        if self.is_pinned and not getattr(self, "was_pinned_before_fs", False):
                            # Auto-Collapse
                            print("DEBUG: Fullscreen App Detected ({}) - Auto Collapsing".format(cls))
                            self.was_pinned_before_fs = True
                            self.is_pinned = False
                            
                            # Update Tooltip/Icon manually since we are bypassing toggle_pin
                            if hasattr(self, 'pin_tooltip'):
                                self.pin_tooltip.text = "Pin Window (Current: Auto-Collapse)"
                            self.draw_pin_icon() 
                            
                            self.save_config() # Optional: Persist? Maybe not if it's temporary state.
                            self.apply_state()
                    
                    else:
                        # Not fullscreen, but on same monitor
                        # Restore if we auto-collapsed
                        if getattr(self, "was_pinned_before_fs", False):
                            # print("DEBUG: Fullscreen ended - Restoring Pin")
                            self.is_pinned = True
                            self.was_pinned_before_fs = False
                            
                            if hasattr(self, 'pin_tooltip'):
                                self.pin_tooltip.text = "Unpin Window (Current: Pinned)"
                            self.draw_pin_icon()
                            
                            self.save_config()
                            self.apply_state()
            
            # If we switched to a different monitor (or desktop focus), we also might want to restore?
            # E.g. User Alt-Tabs to an app on secondary monitor. Sidebar on Primary should probably restore?
            # Current logic only restores if the active window is on SAME monitor and NOT fullscreen.
            # If active window is on OTHER monitor, 'on_same_monitor' is False.
            # We should probably restore if the user leaves the fullscreen app too.
             
            if not on_same_monitor and getattr(self, "was_pinned_before_fs", False):
                 # Focus moved away from the fullscreen app on this monitor
                 # Restore
                 # print("DEBUG: Focus moved monitor - Restoring Pin")
                 self.is_pinned = True
                 self.was_pinned_before_fs = False
                 if hasattr(self, 'pin_tooltip'):
                     self.pin_tooltip.text = "Unpin Window (Current: Pinned)"
                 self.draw_pin_icon()
                 self.save_config()
                 self.apply_state()

        except Exception as e:
            print("FS Check Error: {}".format(e))
            
        self.after(1000, self.check_fullscreen_app)

    def check_updates(self):
        """Threaded (or scheduled) update check."""
        try:
            self._perform_check()
        except Exception as e:
            print("Polling error: {}".format(e))
        
        # Schedule next poll
        interval = getattr(self, "poll_interval", 15) * 1000
        self.after(interval, self.check_updates)

    def _perform_check(self):
        """Actual check logic."""
        accounts = None
        if self.enabled_accounts:
            accounts = list(self.enabled_accounts.keys())

        # 1. Check New Mail (For Refreshing List)
        has_new = self.outlook_client.check_new_mail(accounts)
        if has_new:
             print("DEBUG: Refreshing emails...")
             self.refresh_emails()
        
        # 2. Gather Statuses for Pulse
        unread_count = self.outlook_client.get_unread_count(accounts, self.enabled_accounts)
        due_status = self.outlook_client.get_due_status(accounts) # Returns dict

        active_colors = []
        
        # Priority 1: Unread Mail (Blue)
        if unread_count > 0:
            active_colors.append("#0078D4")

        # Priority 2: Meetings (Orange)
        if due_status["calendar"] == "Today" and self.reminder_show_meetings and "Today" in self.reminder_meeting_dates:
            active_colors.append("#E68D49") # Soft Orange
            
        # Priority 3: Tasks (Green)
        if self.reminder_show_tasks:
            t_status = due_status["tasks"]
            if t_status == "Overdue" and "Overdue" in self.reminder_due_filters:
                active_colors.append("#28C745")
            elif t_status == "Today" and "Today" in self.reminder_due_filters:
                active_colors.append("#28C745")
            
        # Trigger Pulse if needed
        if active_colors and not self.is_pinned and not self.is_expanded:
            print("DEBUG: Active Pulse Colors: {}".format(active_colors))
            self.start_pulse(active_colors)
        elif not active_colors:
            self.stop_pulse()

    def start_pulse(self, colors):
        """Starts the hot strip pulsing animation with a list of colors."""
        if self.is_pinned or self.is_expanded: return 
        
        print("DEBUG: start_pulse triggered. Colors={}".format(colors)) 
        
        # Ensure colors is a list
        if isinstance(colors, str): colors = [colors]
        
        # Update colors if already running
        self.pulse_colors = colors
        
        if not self.pulse_active:
            self.pulse_active = True
            self.pulse_step = 0
            if not getattr(self, "pulse_timer", None):
                self.animate_pulse()

    def stop_pulse(self):
        """Stops the pulsing animation."""
        self.pulse_active = False
        if getattr(self, "pulse_timer", None):
            self.after_cancel(self.pulse_timer)
            self.pulse_timer = None
        
        # Reset color
        if hasattr(self, "hot_strip_canvas"):
            self.hot_strip_canvas.config(bg="#444444")
            self.hot_strip_canvas.delete("pulse_center")

    def animate_pulse(self):
        """Animating loop for pulse: Stacked, Fixed Height, Fading Opacity."""
        if not self.pulse_active or not getattr(self, "pulse_colors", None): 
            self.stop_pulse()
            return
        
        cycle_len = 40
        colors = self.pulse_colors
        num_colors = len(colors)
        
        # Calculate Opacity/Brightness Step (0 to 1.0)
        local_step = self.pulse_step % cycle_len
        # Triangle wave: 0 -> 20 (1.0) -> 40 (0)
        scale = local_step if local_step <= (cycle_len // 2) else (cycle_len - local_step)
        # Normalize to 0.2 - 1.0 range (Never fully invisible)
        brightness = 0.3 + (0.7 * (scale / 20.0)) 
        
        self.hot_strip_canvas.delete("pulse_center")
        
        # Geometry
        item_h = 110 # Fixed height (Reduced from 150)
        w = self.hot_strip_width
        
        # Calculate total stack height
        total_h = num_colors * item_h + ((num_colors - 1) * 12) # Include Gaps
        start_y = (self.winfo_height() // 2) - (total_h // 2)
        
        for i, hex_color in enumerate(colors):
            y1 = start_y + (i * item_h) + (i * 12) # Gap of 12px
            y2 = y1 + item_h
            
            # Interpolate Color for "Fading" effect
            faded_color = self.adjust_color_brightness(hex_color, brightness)
            
            self.hot_strip_canvas.create_rectangle(
                0, y1, w, y2,
                fill=faded_color,
                outline="",
                tags="pulse_center"
            )
        
        self.pulse_step += 1
        # Previous was 15ms. 10% slower = ~17ms.
        self.pulse_timer = self.after(17, self.animate_pulse)

    def adjust_color_brightness(self, hex_color, factor):
        """Dim a hex color by factor (0.0 to 1.0). Simulates opacity over dark bg."""
        # Parse Hex
        if not hex_color.startswith("#"): return hex_color
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        
        # Interpolate towards background (#444444 approx rgb(68,68,68))
        # Or just dim towards black if we want "pulsing off"
        # User said "pulsing off", usually implies fading out. 
        # Fading to black is easiest. Fading to BG #444444 is "opacity".
        bg_r, bg_g, bg_b = 68, 68, 68
        
        # Linear interpolation
        nr = int(bg_r + (r - bg_r) * factor)
        ng = int(bg_g + (g - bg_g) * factor)
        nb = int(bg_b + (b - bg_b) * factor)
        
        # Clamp
        nr = max(0, min(255, nr))
        ng = max(0, min(255, ng))
        nb = max(0, min(255, nb))
        
        return "#{:02x}{:02x}{:02x}".format(nr, ng, nb)

    def on_enter(self, event):
        # Note: We do NOT stop pulsing here anymore. 
        # We want it to keep pulsing until we actually expand.
        
        if self._collapse_timer:
            self.after_cancel(self._collapse_timer)
            self._collapse_timer = None
        
        if not self.is_pinned and not self.is_expanded:
            # Start hover timer (0.75s delay)
            if not self._hover_timer:
                self._hover_timer = self.after(750, self.do_expand)

    def do_expand(self):
        """Actually expands the sidebar after delay."""
        self._hover_timer = None
        if not self.is_pinned and not self.is_expanded:
            self.stop_pulse() # Stop pulse only when genuinely opening
            self.is_expanded = True
            self.apply_state() # Expand and reserve space

    def on_leave(self, event):
        # We need to be careful. Leaving the window to the desktop should collapse.
        # But verify we aren't just hovering a child widget (Tkinter events bubble, but checking coordinates keeps us safe).
        x, y = self.winfo_pointerxy()
        widget_under_mouse = self.winfo_containing(x, y)
        
        # If we are really outside the window
        if not self.is_pinned:
            # Cancel potential expand timer if we left quickly (mouse-over between screens)
            if self._hover_timer:
                self.after_cancel(self._hover_timer)
                self._hover_timer = None
                
            if self.is_expanded:
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
                print("Error creating app data dir: {}".format(e))
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
                        print("Failed to copy default config: {}".format(e))
            
            if os.path.exists(config_path):
                with open(config_path, "r") as f:
                    config = json.load(f)
                    
                self.window_mode = config.get("window_mode", "dual")
                self.split_sash_pos = config.get("split_sash_pos", 0)
                self.enabled_accounts = config.get("enabled_accounts", {})

                self.dock_side = config.get("dock_side", "Right")
                self.font_family = config.get("font_family", "Segoe UI")
                self.font_size = config.get("font_size", 9)
                self.poll_interval = config.get("poll_interval", 15)
                
                # Email List Filters
                self.show_read = config.get("show_read", False)
                self.show_has_attachment = config.get("show_has_attachment", True)

                if "buttons" in config:
                     self.btn_config = config["buttons"]
                     self.btn_count = len(self.btn_config)
                
                self.buttons_on_hover = config.get("buttons_on_hover", True)
                self.email_double_click = config.get("email_double_click", True)
                self.show_hover_content = config.get("show_hover_content", False) # New setting
                
                # Quick Create
                self.quick_create_actions = config.get("quick_create_actions", ["New Email"])
                     
                # Reminder Settings
                self.reminder_show_flagged = config.get("reminder_show_flagged", True)
                self.reminder_due_filters = config.get("reminder_due_filters", ["No Date"])
                
                self.reminder_show_categorized = config.get("reminder_show_categorized", True)
                # self.reminder_categories = config.get("reminder_categories", [])
                
                self.reminder_show_importance = config.get("reminder_show_importance", True)
                self.reminder_high_importance = config.get("reminder_high_importance", False)
                self.reminder_normal_importance = config.get("reminder_normal_importance", False)
                self.reminder_low_importance = config.get("reminder_low_importance", False)
                
                self.reminder_show_meetings = config.get("reminder_show_meetings", True)
                self.reminder_pending_meetings = config.get("reminder_pending_meetings", True)
                self.reminder_accepted_meetings = config.get("reminder_accepted_meetings", True)
                self.reminder_declined_meetings = config.get("reminder_declined_meetings", True)
                self.reminder_meeting_dates = config.get("reminder_meeting_dates", ["Today", "Tomorrow"])
                self.reminder_custom_days = config.get("reminder_custom_days", 30)
                
                self.reminder_show_tasks = config.get("reminder_show_tasks", True)
                self.reminder_tasks = config.get("reminder_tasks", True)
                self.reminder_todo = config.get("reminder_todo", True)
                self.reminder_has_reminder = config.get("reminder_has_reminder", True)
                
                # Email Content Settings
                self.email_show_sender = config.get("email_show_sender", True)
                self.email_show_subject = config.get("email_show_subject", True)
                self.email_show_body = config.get("email_show_body", False)
                self.email_body_lines = config.get("email_body_lines", 2)
        except Exception as e:
            print("Error loading config: {}".format(e))

    def save_config(self):
        app_dir = self.get_app_data_dir()
        config_path = os.path.join(app_dir, "config.json")
        
        config = {
            "dock_side": self.dock_side,
            "window_mode": self.window_mode,
            "enabled_accounts": self.enabled_accounts,
            "split_sash_pos": self.split_sash_pos,
            "font_family": self.font_family,
            "font_size": self.font_size,
            "poll_interval": self.poll_interval,
            "buttons": self.btn_config,
            "buttons_on_hover": self.buttons_on_hover,
            "email_double_click": self.email_double_click,
            
            # Reminder Settings
            "reminder_show_flagged": self.reminder_show_flagged,
            "reminder_due_filters": self.reminder_due_filters,
            
            "reminder_show_categorized": self.reminder_show_categorized,
            # "reminder_categories": self.reminder_categories,
            
            "reminder_show_importance": self.reminder_show_importance,
            "reminder_high_importance": self.reminder_high_importance,
            "reminder_normal_importance": self.reminder_normal_importance,
            "reminder_low_importance": self.reminder_low_importance,
            
            "reminder_show_meetings": self.reminder_show_meetings,
            "reminder_pending_meetings": self.reminder_pending_meetings,
            "reminder_accepted_meetings": self.reminder_accepted_meetings,
            "reminder_declined_meetings": self.reminder_declined_meetings,
            "reminder_meeting_dates": self.reminder_meeting_dates,
            "reminder_custom_days": self.reminder_custom_days,
            
            "reminder_show_tasks": self.reminder_show_tasks,
            "reminder_tasks": self.reminder_tasks,
            "reminder_todo": self.reminder_todo,
            "reminder_has_reminder": self.reminder_has_reminder,
            
            # Email Content Settings
            "email_show_sender": self.email_show_sender,
            "email_show_subject": self.email_show_subject,
            "email_show_body": self.email_show_body,
            "email_body_lines": self.email_body_lines,
            
            # Quick Create
            "quick_create_actions": self.quick_create_actions
        }
        try:
            with open(config_path, "w") as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print("Error saving config: {}".format(e))

    def handle_quick_create(self):
        """Handles the quick create button click."""
        actions = getattr(self, "quick_create_actions", ["New Email"])
        
        # If disabled/empty, do nothing
        if not actions:
             return

        if len(actions) == 1:
            self._execute_quick_action(actions[0])
        else:
            # Show menu
            menu = tk.Menu(self, tearoff=0)
            
            def make_cmd(act):
                return lambda: self._execute_quick_action(act)

            for act in actions:
                menu.add_command(label=act, command=make_cmd(act))
            
            try:
                x = self.btn_quick_create.winfo_rootx()
                y = self.btn_quick_create.winfo_rooty()
                menu.tk_popup(x, y)
            except: pass

    def update_quick_create_icon(self):
        """Updates the Quick Create button icon/color based on selection."""
        if not hasattr(self, "btn_quick_create"): return

        actions = getattr(self, "quick_create_actions", [])
        
        # Determine Color
        color = "#555555" # Grey (Disabled)
        if not actions:
            color = "#555555"
        elif len(actions) > 1:
            color = "#FFFFFF" # White (Multi)
        else:
            # Single Action Colors
            act = actions[0]
            if act == "New Email": color = "#6fb7ff"      # Light Blue
            elif act == "New Meeting": color = "#ffb366"  # Light Orange
            elif act == "New Appointment": color = "#ffcc80" # Lighter Orange
            elif act == "New Task": color = "#80e0a0"     # Light Green
            
        # Update Icon
        try:
            # 26x26 Size
            if os.path.exists("icon2/plus.png"):
                img = self.load_icon_colored("icon2/plus.png", size=(26, 26), color=color)
                self.image_cache["quick_create"] = img
                self.btn_quick_create.configure(image=img)
        except: pass

    def _execute_quick_action(self, action):
        if action == "New Email": self.outlook_client.create_email()
        elif action == "New Meeting": self.outlook_client.create_meeting()
        elif action == "New Appointment": self.outlook_client.create_appointment()
        elif action == "New Task": self.outlook_client.create_task()

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
    try:
        # Check Single Instance
        app_instance = SingleInstance()
        if app_instance.already_running():
            sys.exit(0)

        # Keep the mutex handle alive for the duration of the app
        app = SidebarWindow()
        app.mainloop()
    except Exception as e:
        import traceback
        with open("startup_error.log", "w") as f:
            f.write("Startup Error: {}\n".format(e))
            f.write(traceback.format_exc())

