import tkinter as tk
import ctypes
from ctypes import wintypes
import time
import json
import os
import win32com.client
import re
import math # Added for animation

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

    def set_pos(self, width, screen_width, screen_height):
        """
        Reserves space for the AppBar. 
        Note: If width is small (collapsed), we might effectively reserve 0 space 
        but still keep it registered to handle messaging, or we unregister logic 
        can be handled by the caller.
        """
        if not self.registered:
            return

        # 1. Query Position
        self.abd.rc.left = 0
        self.abd.rc.top = 0
        self.abd.rc.right = width
        self.abd.rc.bottom = screen_height
        
        # Query the system for an approved position
        shell32.SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(self.abd))
        
        # 2. Adjust if necessary (here we force left edge, full height)
        # The system might have adjusted rc, we re-apply our desired width
        # constrained by what the system gave us.
        self.abd.rc.right = self.abd.rc.left + width
        
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
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=kwargs.get("bg", "#222222"))

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Ensure scrollable frame matches canvas width
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def _on_canvas_configure(self, event):
        # Resize the inner frame to match the canvas width
        self.canvas.itemconfig(self.window_id, width=event.width)

class OutlookClient:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.last_received_time = None
        self.connect()
        # Initialize last_received_time
        self.check_latest_time()

    def connect(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            print(f"Error connecting to Outlook: {e}")

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
        except:
            pass

    def check_new_mail(self):
        """Checks if there is email newer than the last check."""
        if not self.namespace: return False
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
        except Exception as e:
            print(f"Polling error: {e}")
        return False

    def get_inbox_items(self, count=20):
        if not self.namespace:
            return []
        
        try:
            inbox = self.namespace.GetDefaultFolder(6) # 6 = olFolderInbox
            items = inbox.Items
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
                        "unread": unread
                    })
                except Exception as inner_e:
                    print(f"Error reading item: {inner_e}")
                    
            return email_list
        except Exception as e:
            print(f"Error fetching items: {e}")
            return []

class SidebarWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- Configuration ---
        self.min_width = 80  
        self.hot_strip_width = 10
        self.expanded_width = 300
        self.is_pinned = False
        self.is_expanded = False
        self.hover_delay = 500 # ms
        self._hover_timer = None
        self._collapse_timer = None
        
        # Pulse Animation State
        self.pulsing = False
        self.pulse_step = 0
        self._pulse_job = None
        self.animation_speed = 0.05 # Increment per frame
        self.base_color = "#007ACC"
        self.pulse_color = "#99D9EA" # Lighter cyan/blue for the bar
        
        # Load Config
        self.load_config()

        # --- Outlook Client ---
        self.outlook_client = OutlookClient()

        # --- Window Setup ---
        self.overrideredirect(True)  # Frameless
        self.wm_attributes("-topmost", True)
        self.config(bg="#333333")

        # Get Screen Dimensions
        self.screen_width = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()

        # --- AppBar Manager ---
        self.update_idletasks() 
        self.hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
        if not self.hwnd:
             self.hwnd = self.winfo_id()

        self.appbar = AppBarManager(self.hwnd)
        
        # --- UI Components ---
        self.main_frame = tk.Frame(self, bg="#222222")
        self.main_frame.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Header
        self.header = tk.Frame(self.main_frame, bg="#444444", height=40)
        self.header.pack(fill="x", side="top")
        
        # Title
        self.lbl_title = tk.Label(self.header, text="Outlook Inbox", bg="#444444", fg="white", font=("Segoe UI", 10, "bold"))
        self.lbl_title.pack(side="left", padx=10)

        # Pin Button / Logo (Custom Canvas)
        self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg="#444444", highlightthickness=0)
        self.btn_pin.pack(side="right", padx=5, pady=5)
        self.btn_pin.bind("<Button-1>", lambda e: self.toggle_pin())
        self.draw_pin_icon()
        
        # Refresh Button
        self.btn_refresh = tk.Button(self.header, text="↻", command=self.refresh_emails, bg="#555555", fg="white", bd=0)
        self.btn_refresh.pack(side="right", padx=5)

        # Content Area - Scrollable Frame for Emails
        self.content_container = tk.Frame(self.main_frame, bg="#222222")
        self.content_container.pack(expand=True, fill="both", padx=5, pady=5)
        
        self.scroll_frame = ScrollableFrame(self.content_container, bg="#222222")
        self.scroll_frame.pack(expand=True, fill="both")

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
        
        # Start Polling
        self.start_polling()
        
    def start_polling(self):
        """Poll Outlook every 30 seconds for new mail."""
        if self.outlook_client.check_new_mail():
            self.start_pulse()
            self.refresh_emails() # Auto-refresh list
            
        self.after(30000, self.start_polling) # 30s
        
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
        
    def load_config(self):
        try:
            with open("sidebar_config.json", "r") as f:
                data = json.load(f)
                self.expanded_width = data.get("width", 300)
                self.is_pinned = data.get("pinned", False)
        except FileNotFoundError:
            pass

    def save_config(self):
        data = {
            "width": self.expanded_width,
            "pinned": self.is_pinned
        }
        with open("sidebar_config.json", "w") as f:
            json.dump(data, f)

    def refresh_emails(self):
        # Clear existing
        for widget in self.scroll_frame.scrollable_frame.winfo_children():
            widget.destroy()

        emails = self.outlook_client.get_inbox_items(count=30)
        
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
            
            # Sender
            sender_text = email['sender']
            if is_unread:
                sender_text = "● " + sender_text # Add indicator dot
                
            lbl_sender = tk.Label(
                card, 
                text=sender_text, 
                fg="white", 
                bg=bg_color, 
                font=("Segoe UI", 9, "bold"),
                anchor="w"
            )
            lbl_sender.pack(fill="x")
            
            # Subject
            lbl_subject = tk.Label(
                card, 
                text=email['subject'], 
                fg="#cccccc", 
                bg=bg_color, 
                font=("Segoe UI", 9),
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
                font=("Segoe UI", 8),
                anchor="w",
                justify="left",
                wraplength=self.expanded_width - 40 
            )
            lbl_preview.pack(fill="x")
            
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
        if self.is_pinned:
            # Pinned: Always Expanded, Always Reserved (Docked)
            self.hot_strip_canvas.place_forget()
            self.header.pack(fill="x", side="top")
            self.content_container.pack(expand=True, fill="both", padx=5, pady=5)
            self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
            
            self.set_geometry(self.expanded_width)
            self.appbar.register()
            self.appbar.set_pos(self.expanded_width, self.screen_width, self.screen_height)
            self.is_expanded = True
            
        elif self.is_expanded:
            # Expanded (Hover): Broad width, BUT acts as OVERLAY (No docking/reservation)
            self.hot_strip_canvas.place_forget()
            self.header.pack(fill="x", side="top")
            # For overlay mode, we still show the content
            self.content_container.pack(expand=True, fill="both", padx=5, pady=5)
            self.resize_grip.place(relx=1.0, rely=0, anchor="ne", relheight=1.0)
            
            # Unregister AppBar so we don't push other windows
            self.appbar.unregister()
            
            self.set_geometry(self.expanded_width)
            
        else:
            # Collapsed: Thin width, Overlay
            self.appbar.unregister() # Release space
            
            # Hide internals to prevent squishing
            self.header.pack_forget()
            self.content_container.pack_forget()
            self.resize_grip.place_forget()
            
            # Show Hot Strip
            self.hot_strip_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
            
            self.set_geometry(self.hot_strip_width)

    def on_resize_drag(self, event):
        if self.is_pinned or self.is_expanded:
            x_root = self.winfo_pointerx()
            new_width = x_root
            
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

        # Always dock left, full height
        self.geometry(f"{width}x{self.screen_height}+0+0")
        # Force top most again just in case
        self.wm_attributes("-topmost", True)

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

if __name__ == "__main__":
    app = SidebarWindow()
    app.mainloop()
