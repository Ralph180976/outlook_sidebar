# -*- coding: utf-8 -*-
import ctypes
from ctypes import wintypes

# --- DPI Awareness (must run BEFORE tkinter creates any windows) ---
def set_dpi_awareness():
    try:
        user32 = ctypes.windll.user32
    except:
        return # Non-Windows or other issue

    # Prefer Per-Monitor v2 on Win10+ (best coordinate correctness across mixed-DPI monitors)
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetProcessDpiAwarenessContext.restype = ctypes.c_int # BOOL
        
        ctx = ctypes.c_void_p(-4)
        if user32.SetProcessDpiAwarenessContext(ctx):
            return
    except Exception:
        pass 

    # Fallback: Win 8.1+ system DPI aware
    try:
        shcore = ctypes.windll.shcore
        # PROCESS_SYSTEM_DPI_AWARE = 1
        shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        pass

    # Last resort: Vista+ system DPI aware
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

set_dpi_awareness()

try:
    # Python 2
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
    import tkFileDialog as filedialog
except ImportError:
    # Python 3
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

import sys
import os
import json
import time
import math
import glob
import ctypes

try:
    import sentry_sdk
    sentry_sdk.init(
        dsn="https://3542f3c3d42e6dea7747f1a9ae88af18@o4510942810472448.ingest.de.sentry.io/4510945291010128",
        send_default_pii=True,
    )
except ImportError:
    pass

from ctypes import wintypes
from PIL import Image, ImageTk, ImageDraw
from datetime import datetime, timedelta
import win32gui
import win32con

kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

if not hasattr(wintypes, 'HMONITOR'):
    wintypes.HMONITOR = wintypes.HANDLE

MONITORENUMPROC = ctypes.WINFUNCTYPE(     ctypes.c_int,      wintypes.HMONITOR,      wintypes.HDC,      ctypes.POINTER(wintypes.RECT),      wintypes.LPARAM )

# --- Modular Imports ---
from sidebar.core.config import (
    VERSION, RESAMPLE_MODE, DEFAULT_MIN_WIDTH, 
    DEFAULT_HOT_STRIP_WIDTH, DEFAULT_EXPANDED_WIDTH, 
    DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE,
    resource_path
)
from sidebar.core.config_manager import ConfigManager
from sidebar.core.theme import COLOR_PALETTES, OL_CAT_COLORS
from sidebar.core.appbar import AppBarManager, MONITORINFO, ABE_LEFT, ABE_RIGHT, ABE_TOP, ABE_BOTTOM 
from sidebar.services.outlook_client import OutlookClient
from sidebar.services.graph_client import GraphAPIClient
from sidebar.services.hybrid_client import HybridMailClient
from sidebar.ui.widgets.base import ScrollableFrame, RoundedFrame, ToolTip
from sidebar.ui.panels.settings import SettingsPanel
from sidebar.ui.panels.help import HelpPanel
from sidebar.ui.panels.account_settings import AccountSelectionDialog, AccountSelectionUI, FolderPickerFrame
from sidebar.ui.panels.account_settings import AccountSelectionDialog, AccountSelectionUI, FolderPickerFrame
from sidebar.ui.dialogs.feedback import FeedbackDialog
from sidebar.ui.widgets.toolbar import SidebarToolbar
from sidebar.services.update_checker import check_for_update

class SidebarWindow(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        # --- Configuration Manager ---
        self.config = ConfigManager()
        
        # Shortcuts for compatibility during refactor (properties wrapping config)
        # Or just use self.config.x directly.
        
        # --- Theme & Colors ---
        try:
            self.current_theme = self.config.theme
            self.font_family = self.config.font_family
        except Exception as e:
            print("ERROR: Failed to access theme: {}".format(e))
            self.current_theme = "Light"

        self.palettes = {
            "Dark": {
                "bg_root": "#333333",
                "bg_header": "#444444",
                "bg_card": "#252526",
                "bg_card_hover": "#2a2a2d",
                "bg_task": "#2d2d2d",
                "fg_primary": "#FFFFFF",
                "fg_text": "#FFFFFF", # Alias for SettingsPanel compatibility
                "fg_secondary": "#CCCCCC",
                "fg_dim": "#999999",
                "accent": "#60CDFF",
                "divider": "#3E3E42",
                "scroll_bg": "#222222",
                "input_bg": "#1E1E1E",
                "card_border": "#333333"
            },
            "Light": {
                "bg_root": "#F3F3F3",
                "bg_header": "#E0E0E0",
                "bg_card": "#FFFFFF",
                "bg_card_hover": "#F9F9F9",
                "bg_task": "#FFFFFF",
                "fg_primary": "#000000",
                "fg_text": "#000000", # Alias
                "fg_secondary": "#333333",
                "fg_dim": "#666666",
                "accent": "#3AADE5",
                "divider": "#D0D0D0",
                "scroll_bg": "#E8E8E8",
                "input_bg": "#FFFFFF",
                "card_border": "#E5E5E5"
            }
        }
        self.colors = self.palettes.get(self.current_theme, self.palettes["Light"])

        self._hover_timer = None
        self._collapse_timer = None
        self.is_expanded = False
        self.hot_strip_width = DEFAULT_HOT_STRIP_WIDTH
        
        # Settings Panel State
        self.settings_panel_open = False
        self.settings_panel = None
        self.settings_panel_width = 370
        
        # Window Mode State
        # self.split_sash_pos = 0 # Now in self.config
        
        # Pulse Animation State
        self.pulsing = False
        self.pulse_step = 0
        self._pulse_job = None
        self.animation_speed = 0.05 # Increment per frame
        self.base_color = "#007ACC"
        self.pulse_color = "#99D9EA" # Lighter cyan/blue for the bar
        
        # self.show_hover_content = False # Now in self.config
        # self.quick_create_actions = ["New Email"] # Now in self.config
        
        self.help_panel = None
        self.help_panel_open = False
        
        try:
            self.outlook_client = self._select_backend()
        except Exception as e:
            print("ERROR: MailClient init failed: {}".format(e))
            import traceback
            traceback.print_exc()
            try:
                import sentry_sdk
                sentry_sdk.capture_exception()
            except ImportError:
                pass
            self.outlook_client = None
        
        # Image Cache (to keep references alive)
        self.image_cache = {}
        self.dismissed_calendar_ids = set()  # Track dismissed calendar items (session only)
        self._calendar_widgets = []  # [(start_dt, time_label, subj_label, frame)] for urgency updates
        self._cal_urgency_timer = None  # Timer for periodic urgency checks

        # --- Window Setup ---
        self.overrideredirect(True)  # Frameless
        self.wm_attributes("-topmost", True)
        self.config_window_visuals()

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
        try:
             self.appbar.hook_wndproc()
        except Exception as e:
             print("Failed to hook WndProc: {}".format(e))
        
        # --- UI Components ---
        # Container frame that holds main content and settings panel side by side
        self.content_wrapper = tk.Frame(self, bg=self.colors["bg_root"])
        self.content_wrapper.pack(fill="both", expand=True)
        
        # Main sidebar content frame (expands to fill space when settings closed)
        self.main_frame = tk.Frame(self.content_wrapper, bg=self.colors["bg_root"])
        self.main_frame.pack(side="left", fill="both", expand=True)

        # Footer
        self.footer = tk.Frame(self.main_frame, bg=self.colors["bg_header"], height=40)
        self.footer.pack(fill="x", side="bottom")
        
        # Header
        self.header = tk.Frame(self.main_frame, bg=self.colors["bg_header"], height=40)
        self.header.pack(fill="x", side="top")
        
        self.header.bind("<Button-1>", self.start_window_drag)
        self.header.bind("<B1-Motion>", self.on_window_drag)
        self.header.bind("<ButtonRelease-1>", self.stop_window_drag)
        
        # Title
        self.lbl_title = tk.Label(self.header, text="InboxBar", bg=self.colors["bg_header"], fg=self.colors["fg_primary"], font=(self.font_family, 10, "bold"))
        self.lbl_title.pack(side="left", padx=10)
        self.lbl_title.bind("<Button-1>", self.start_window_drag)
        self.lbl_title.bind("<B1-Motion>", self.on_window_drag)
        self.lbl_title.bind("<ButtonRelease-1>", self.stop_window_drag)

        # Pin Button / Logo (Custom Canvas)
        # Initialize Toolbar
        # ------------------
        callbacks = {
            "settings": self.open_settings,
            "help": self.toggle_help_panel,
            "refresh": self.refresh_emails,
            "share": self.open_share_dialog,
            "close": self.quit_application,
            "quick_create": self.handle_quick_create,
            "calendar": self.open_calendar_app,
            "outlook": self.open_outlook_app,
            "toggle_pin": self.toggle_pin
        }
        
        self.toolbar = SidebarToolbar(
            self.header, self.footer, callbacks, 
            self.load_icon_colored, resource_path, self.config
        )
        self.toolbar.create_header_buttons(self.colors)
        self.toolbar.create_footer_buttons(self.colors, version_text=VERSION)

        # Proxies for external access (if any legacy code tries to access buttons directly)
        # Ideally we remove these, but for safety in Phase 2 we can alias them
        self.btn_pin = self.toolbar.btn_pin
        self.btn_settings = self.toolbar.btn_settings
        self.btn_help = self.toolbar.btn_help
        self.btn_refresh = self.toolbar.btn_refresh
        self.btn_share = self.toolbar.btn_share
        self.btn_outlook = self.toolbar.btn_outlook
        self.btn_calendar = self.toolbar.btn_calendar
        self.btn_quick_create = self.toolbar.btn_quick_create
        self.btn_close = self.toolbar.btn_close
        self.lbl_version = self.toolbar.lbl_version

        # Content Area - Using PanedWindow for draggable resizing
        # Replaces grid_container
        
        # PanedWindow
        # sashwidth=4, sashrelief="raised" or "flat", bg="#333333" for visibility
        # opaqueresize=False preventing constant redraw during drag (smoother)
        # PanedWindow
        # sashwidth=4, sashrelief="raised" or "flat", bg="bg_root" for visibility
        # opaqueresize=False preventing constant redraw during drag (smoother)
        self.paned_window = tk.PanedWindow(self.main_frame, orient="vertical", bg=self.colors["bg_root"], sashwidth=8, sashrelief="flat", opaqueresize=False)
        self.paned_window.pack(expand=True, fill="both", padx=5, pady=5)
        
        # Pane 1: Email Container
        self.pane_emails = tk.Frame(self.paned_window, bg=self.colors["bg_root"])
        self.paned_window.add(self.pane_emails, minsize=100)
        
        # Pane 2: Reminder Container
        self.pane_reminders = tk.Frame(self.paned_window, bg=self.colors["bg_root"])
        self.paned_window.add(self.pane_reminders, minsize=100)
        
        # Note: We need to set the sash position AFTER the window is rendered/updated.
        # Set 50/50 split after geometry is available
        def set_initial_sash():
            try:
                self.paned_window.update_idletasks()
                h = self.paned_window.winfo_height()
                if h > 50:
                    self.paned_window.sash_place(0, 0, h // 2)
            except:
                pass
        self.after(300, set_initial_sash)
        
        # Email section header (Created once, kept in pane_emails)
        self.email_header = tk.Frame(self.pane_emails, bg=self.colors["bg_card"], height=26)
        self.email_header.pack(fill="x", side="top")
        self.email_header.pack_propagate(False)  # Maintain fixed height
        
        # Email List Container (Scrollable) - Will be filled in refresh_emails
        self.email_list_frame = tk.Frame(self.pane_emails, bg=self.colors["bg_root"])
        self.email_list_frame.pack(fill="both", expand=True)

        # ------------------
        # Header Controls
        # ------------------
        
        self.lbl_email_header = tk.Label(
            self.email_header, text="Email", 
            bg=self.colors["bg_card"], fg=self.colors["fg_text"],
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
        self.btn_account_toggle = tk.Label(self.email_header, cursor="hand2")
        self._update_arrow_icons()
             
        self.btn_account_toggle.pack(side="right", padx=5)
        self.btn_account_toggle.bind("<Button-1>", lambda e: self.toggle_account_selection())
            

        
        # ------------------

        
        self.scroll_frame = ScrollableFrame(self.email_list_frame, bg=self.colors["bg_root"])
        self.scroll_frame.pack(expand=True, fill="both")
        
        # Reminder List Setup (Initial empty state, populated in refresh_reminders)
        # Note: We already added pane_reminders to PanedWindow
        
        # Header inside pane_reminders (for consistency with email pane)
        self.r_header = tk.Frame(self.pane_reminders, bg=self.colors["bg_card"], height=20)
        self.r_header.pack(fill="x", side="top")
        self.r_header.pack_propagate(False)
        
        self.lbl_reminder_header = tk.Label(self.r_header, text="Flagged/Reminders", 
                 bg=self.colors["bg_card"], fg=self.colors["fg_dim"], font=(self.font_family, 9, "bold")
        ).pack(side="left", padx=10, pady=3)

        self.reminder_list = ScrollableFrame(self.pane_reminders, bg=self.colors["input_bg"])
        self.reminder_list.pack(fill="both", expand=True)
        
        # Reminder placeholder removed
        pass

        # Resize Grip (Overlay on the right edge)
        # Resize Grip (Overlay on the right edge)
        self.resize_grip = tk.Frame(self.main_frame, bg=self.colors["fg_dim"], cursor="sb_h_double_arrow", width=5)
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
        
        # Check for updates (runs in background thread)
        self._check_for_app_update()

    def config_window_visuals(self):
        """Configures window visual properties (e.g. transparency)."""
        # Placeholder restored after accidental deletion
        pass

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
            self.set_geometry(self.config.width)
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
                new_width = self.config.width + self.settings_panel.panel_width
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

    def open_share_dialog(self):
        """Open the Share dialog to email the InboxBar installer."""
        from sidebar.ui.dialogs.share import ShareDialog
        ShareDialog(self.winfo_toplevel(), self.outlook_client)

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
            self.set_geometry(self.config.width)
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
            new_width = self.config.width + self.help_panel.panel_width
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
            if "A" in pil_img.getbands():
                 r, g, b, a = pil_img.split()
            else:
                 # Assume opaque if no alpha? Or convert created one.
                 # maximize alpha
                 a = Image.new("L", pil_img.size, 255)
            
            # Create final image: Transparent background + Colored Shape masked by Alpha
            final_img = Image.new("RGBA", pil_img.size, (0, 0, 0, 0))
            final_img.paste(colored_img, (0, 0), mask=a)
            
            return ImageTk.PhotoImage(final_img, master=self)
        except Exception as e:
            print("Error loading/coloring icon {}: {}".format(path, e))
            return None

    def load_icon_white(self, path, size=None):
        """Legacy wrapper for load_icon_colored (defaults to standard grey)."""
        return self.load_icon_colored(path, size, color="#BFBFBF")

    def handle_custom_action(self, config, email_data, source_card=None):
        """Executes the selected actions on the specific email."""
        print("Executing Actions for {} on {}".format(config.get('label'), email_data.get('subject')))
        
        entry_id = email_data.get("entry_id")
        store_id = email_data.get("store_id") # Support multi-account
        if not entry_id:
            print("No EntryID found.")
            return

        # --- Instant visual feedback: remove the card immediately ---
        act1 = config.get("action1", "")
        # Determine which actions actually remove the card from view
        always_removes = act1 in ("Delete", "Read & Delete", "Move To...")
        # Mark Read / Flag only remove from unread-only view
        conditional_removes = act1 in ("Mark Read", "Flag") and not self.config.show_read
        if (always_removes or conditional_removes) and source_card:
            try:
                source_card.pack_forget()
                source_card.destroy()
            except: pass
            # Update header count immediately (for unread counter)
            if act1 != "Flag":
                try:
                    current = int(self.lbl_email_header.cget("text").split(" - ")[1])
                    self.lbl_email_header.config(text="Email - {}".format(max(0, current - 1)))
                except: pass

        # Use MailClient abstraction instead of raw COM objects
        def execute_single_action(act_name, folder_name=""):
            if not act_name or act_name == "None": return
            
            with open("C:\\Dev\\Outlook_Sidebar\\debug_out.txt", "a") as f:
                 f.write(f"Executing action '{act_name}' on ID {entry_id[:15]}...\n")
                 
            try:
                if act_name == "Mark Read":
                    self.outlook_client.mark_as_read(entry_id, store_id)
                elif act_name == "Delete":
                    self.outlook_client.delete_email(entry_id, store_id)
                elif act_name == "Read & Delete":
                    self.outlook_client.mark_as_read(entry_id, store_id)
                    self.outlook_client.delete_email(entry_id, store_id)
                elif act_name == "Flag":
                    self.outlook_client.toggle_flag(entry_id, store_id)
                elif act_name == "Open Email":
                    self._allow_foreground_for_outlook()
                    self.outlook_client.open_item(entry_id, store_id)
                elif act_name == "Reply":
                    # Mark as read first
                    self.outlook_client.mark_as_read(entry_id, store_id)
                    
                    self._allow_foreground_for_outlook()
                    if hasattr(self.outlook_client, "reply_to_email"):
                        self.outlook_client.reply_to_email(entry_id, store_id)
                    else:
                        print("Reply action not yet fully abstracted for Graph API")
                        
                elif act_name == "Move To...":
                    if folder_name:
                         if hasattr(self.outlook_client, "move_email"):
                             self.outlook_client.move_email(entry_id, folder_name, store_id)
                         else:
                             print("Move to Folder not fully abstracted for Graph API")
            except Exception as e:
                print("Error executing {}: {}".format(act_name, e))

        try:
            # Execute Action 1
            execute_single_action(config.get("action1"), config.get("folder"))
            
            # Execute Action 2 - REMOVED
            # execute_single_action(config.get("action2"), config.get("folder"))
                
            # Refresh UI — fast delay since card is already hidden
            # Flag actions need reminders refreshed too
            if act1 == "Flag":
                self.after(100, self.refresh_emails)
            else:
                self.after(100, lambda: self.refresh_emails(skip_reminders=True))
            
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
        Reuses the existing OutlookClient's connection.
        """
        # First, try to reuse the backend (if it's COM based)
        if hasattr(self, 'outlook_client') and self.outlook_client:
            app = self.outlook_client.get_native_app()
            if app: return app
            if hasattr(self.outlook_client, 'connect'):
                 self.outlook_client.connect()
                 return self.outlook_client.get_native_app()
        
        # Fallback: try direct COM (requires win32com import)
        try:
            import win32com.client
            app = win32com.client.GetActiveObject("Outlook.Application")
            return app
        except:
            pass
        
        try:
            import win32com.client
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

    def _allow_foreground_for_outlook(self):
        """Grant our process permission to set the Outlook window as foreground."""
        try:
            import subprocess
            # Find Outlook's PID
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, timeout=3
            )
            for line in result.stdout.strip().split("\n"):
                parts = line.strip().strip('"').split('","')
                if len(parts) >= 2:
                    try:
                        pid = int(parts[1])
                        ctypes.windll.user32.AllowSetForegroundWindow(pid)
                        return
                    except (ValueError, Exception):
                        pass
            # Fallback: allow any process
            ctypes.windll.user32.AllowSetForegroundWindow(-1)  # ASFW_ANY
        except Exception:
            try:
                ctypes.windll.user32.AllowSetForegroundWindow(-1)
            except Exception:
                pass

    def _show_outlook_folder(self, folder_id):
        """
        Shows the specified Outlook folder (6=Inbox, 9=Calendar).
        Reuses existing Explorer if available, otherwise creates one.
        Forces the window to the foreground.
        Returns True if it succeeded, False otherwise.
        """
        import traceback
        succeeded = False
        try:
            # Grant foreground permission before doing anything
            self._allow_foreground_for_outlook()
            
            app = self._get_outlook_app()
            if not app:
                print("[Outlook] Could not get Outlook Application object")
                return False
            
            ns = app.GetNamespace("MAPI")
            folder = ns.GetDefaultFolder(folder_id)
            print("[Outlook] Got folder for id={}".format(folder_id))
            
            # Try to get existing explorer
            explorer = self._get_any_explorer(app)
            
            if explorer:
                print("[Outlook] Found existing explorer")
                # Reuse existing explorer - switch folder
                try:
                    explorer.CurrentFolder = folder
                    print("[Outlook] Switched to folder")
                except Exception as e:
                    print("[Outlook] Failed to switch folder: {}".format(e))
                
                # Activate via COM first
                try:
                    explorer.Activate()
                    print("[Outlook] Activated explorer")
                    succeeded = True
                except Exception as e:
                    print("[Outlook] Activate failed: {}".format(e))
                
                # Then force focus via hwnd
                hwnd = None
                try:
                    hwnd = explorer.Hwnd if hasattr(explorer, 'Hwnd') else None
                    print("[Outlook] Explorer hwnd={}".format(hwnd))
                except Exception as e:
                    print("[Outlook] Failed to get hwnd: {}".format(e))
                    
                if hwnd:
                    self._focus_window_by_hwnd(hwnd)
                    succeeded = True
                    # Delayed re-focus as safety net
                    self.after(150, lambda h=hwnd: self._focus_window_by_hwnd(h))
            else:
                print("[Outlook] No existing explorer found, creating new one")
                # No explorer exists - create one via GetExplorer
                try:
                    new_explorer = folder.GetExplorer()
                    new_explorer.Display()
                    succeeded = True
                    print("[Outlook] Created and displayed new explorer")
                    
                    # Focus the new window with retries
                    def _delayed_focus(attempt=1):
                        if attempt > 5:
                            return
                        try:
                            h = new_explorer.Hwnd if hasattr(new_explorer, 'Hwnd') else None
                            if h:
                                self._focus_window_by_hwnd(h)
                            else:
                                self.after(100, lambda: _delayed_focus(attempt + 1))
                        except Exception:
                            self.after(100, lambda: _delayed_focus(attempt + 1))
                    
                    self.after(100, _delayed_focus)
                except Exception as e:
                    print("[Outlook] GetExplorer failed: {}".format(e))
                    # Ultimate fallback
                    try:
                        folder.Display()
                        succeeded = True
                    except Exception as e2:
                        print("[Outlook] folder.Display() also failed: {}".format(e2))
                    
        except Exception as e:
            print("[Outlook] _show_outlook_folder error: {}".format(e))
            traceback.print_exc()
        
        return succeeded

    def open_outlook_app(self):
        """Opens/Focuses the main Outlook window (Inbox)."""
        success = self._show_outlook_folder(6)  # 6 = olFolderInbox
        if not success:
            # Fallback: launch Outlook via subprocess to show Inbox
            print("[Outlook] COM method failed, falling back to subprocess")
            try:
                import subprocess
                subprocess.Popen(["outlook.exe", "/select", "outlook:inbox"])
            except Exception as e:
                print("[Outlook] subprocess fallback also failed: {}".format(e))
                try:
                    os.startfile("outlook:")
                except Exception as e2:
                    print("[Outlook] os.startfile fallback also failed: {}".format(e2))

    def open_calendar_app(self):
        """Opens/Focuses the Outlook Calendar."""
        success = self._show_outlook_folder(9)  # 9 = olFolderCalendar
        if not success:
            print("[Outlook] COM method failed for calendar, falling back to subprocess")
            try:
                import subprocess
                subprocess.Popen(["outlook.exe", "/select", "outlook:calendar"])
            except Exception as e:
                print("[Outlook] subprocess calendar fallback failed: {}".format(e))
        
        
    # Legacy load_config/save_config removed. 
    # They are now handled by ConfigManager (self.config).


    def _select_backend(self):
        """Instantiates the correct MailClient backend based on config."""
        backend_pref = getattr(self.config, "backend", "hybrid")
        
        if backend_pref == "graph":
            print("[Backend] Using Microsoft Graph API exclusively")
            return GraphAPIClient()
        elif backend_pref == "com":
            print("[Backend] Using Classic Outlook COM exclusively")
            return OutlookClient()
        else: # auto / hybrid
            print("[Backend] Using Hybrid Client (COM + Graph)")
            return HybridMailClient()

    def apply_window_layout(self):
        """Apply the current window mode (single or dual) to the layout."""
        if self.config.window_mode == "single":
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

    def open_email(self, entry_id, source_widget=None, store_id=None, fallback_link=None):
        """Opens the specific email item with visual feedback."""
        if source_widget:
            self.flash_widget_recursive(source_widget)
                
        import string
        is_hex = all(c in string.hexdigits.upper() + string.hexdigits.lower() for c in entry_id) if hasattr(entry_id, 'isalnum') else False
        
        # If it doesn't look like a standard COM hex entry ID
        if not is_hex or len(entry_id) < 40 or "AAMk" in entry_id:
            if fallback_link:
                import webbrowser
                webbrowser.open(fallback_link)
                return
            try:
                self.outlook_client.open_item(entry_id, store_id)
            except Exception as e:
                print("Error routing open to hybrid client: {}".format(e))
            return
            
        try:
             # Safe try for raw COM object interaction with native focus forcing
             if not self.outlook_client.com.namespace:
                 self.outlook_client.com.connect()
             
             if self.outlook_client.com.namespace:
                 self._allow_foreground_for_outlook()
                 if store_id:
                     item = self.outlook_client.com.namespace.GetItemFromID(entry_id, store_id)
                 else:
                     item = self.outlook_client.com.namespace.GetItemFromID(entry_id)
                 item.Display()
                 
                 try:
                     inspector = item.GetInspector
                     inspector.Activate()
                     
                     # Force window usage if possible
                     try:
                        caption = inspector.Caption
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

    def _update_arrow_icons(self):
        try:
            color = self.colors["fg_text"]
            # DOWN Arrow (Closed)
            img_down = Image.new("RGBA", (20, 20), (0,0,0,0))
            draw_d = ImageDraw.Draw(img_down)
            draw_d.polygon([(4, 7), (16, 7), (10, 15)], fill=color)
            self.icon_arrow_down = ImageTk.PhotoImage(img_down)
            
            # UP Arrow (Open)
            img_up = Image.new("RGBA", (20, 20), (0,0,0,0))
            draw_u = ImageDraw.Draw(img_up)
            draw_u.polygon([(4, 15), (16, 15), (10, 7)], fill=color)
            self.icon_arrow_up = ImageTk.PhotoImage(img_up)

            if hasattr(self, "account_overlay") and self.account_overlay and self.account_overlay.winfo_exists():
                self.btn_account_toggle.config(image=self.icon_arrow_up, bg=self.colors["bg_card"])
            else:
                self.btn_account_toggle.config(image=self.icon_arrow_down, bg=self.colors["bg_card"])
        except Exception as e:
            print("Error generating arrow icons: {}".format(e))
            if hasattr(self, "btn_account_toggle") and self.btn_account_toggle:
                self.btn_account_toggle.config(text="▼", bg=self.colors["bg_card"], fg=self.colors["fg_text"])

    def toggle_account_selection(self):
        """Toggles the account selection overlay."""
        try:
            if hasattr(self, "account_overlay") and self.account_overlay and self.account_overlay.winfo_exists():
                # Closing: Save and Destroy
                if hasattr(self, "account_ui_helper"):
                    new_settings = self.account_ui_helper.get_settings()
                    self.config.enabled_accounts = new_settings
                    self.config.save()
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
                # Do not error out if empty, allow UI to show "Add Account" button

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
                    self.config.enabled_accounts, 
                    self.launch_folder_selection_from_overlay, 
                    bg_color="#202020"
                )
                self.account_ui_helper.pack(fill="both", expand=True)
                
                # Raise to top
                self.account_overlay.lift()
        except Exception as e:
            import traceback
            with open("runtime_error.log", "a") as f:
                f.write("Error in toggle_account_selection: {}\\n".format(e))
                f.write(traceback.format_exc())
            messagebox.showerror("Error", "Failed to open account selection:\\n{}".format(e))

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

    def refresh_emails(self, skip_reminders=False):
        if not self.outlook_client: return
        try:
            # Update UI fonts for header elements
            self.lbl_title.config(font=(self.font_family, 10, "bold"))
            self.btn_settings.config(font=(self.font_family, 12))
            self.btn_refresh.config(font=(self.font_family, 15))

            # --- Anti-flicker: hide canvas content during rebuild ---
            canvas = self.scroll_frame.canvas
            canvas.itemconfigure(self.scroll_frame.window_id, state='hidden')

            # Clear existing
            for widget in self.scroll_frame.scrollable_frame.winfo_children():
                widget.destroy()

            # Determine enabled accounts
            accounts = [n for n, s in self.config.enabled_accounts.items() if s.get("email")] if self.config.enabled_accounts else None


            emails, unread_count = self.outlook_client.get_inbox_items(

                count=30, 
                unread_only=not self.config.show_read,
                account_names=accounts,
                account_config=self.config.enabled_accounts
            )
            
            # Update Header Count
            try:
                 self.lbl_email_header.config(text="Email - {}".format(unread_count), bg=self.colors["bg_card"], fg=self.colors["fg_text"])
            except: pass
            
            # Fetch Category Colors (cached with 5-min TTL)
            now_ts = time.time()
            if not hasattr(self, '_cat_map_cache') or now_ts - getattr(self, '_cat_map_cache_time', 0) > 300:
                self._cat_map_cache = self.outlook_client.get_category_map()
                self._cat_map_cache_time = now_ts
            cat_map = self._cat_map_cache
            
            for email in emails:
                lbl_sender = None
                lbl_subject = None
                lbl_preview = None
                
                # Determine styling based on UnRead status
                is_unread = email.get('unread', False)
                bg_color = self.colors["bg_card"]
                # Blue border for unread, grey for read
                border_color = self.colors["accent"] if is_unread else self.colors["card_border"]
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
                if self.config.email_show_sender:
                    sender_text = email['sender']
                    if is_unread:
                        sender_text = u"● " + sender_text # Add indicator dot
                        
                    lbl_sender = tk.Label(
                        header_frame, 
                        text=sender_text, 
                        fg=self.colors["fg_primary"], 
                        bg=bg_color, 
                        font=(self.config.font_family, self.config.font_size, "bold"),
                        anchor="w"
                    )
                    lbl_sender.pack(side="left", fill="x", expand=True)


                # Attachment indicator (only show if setting is enabled)
                if email.get('has_attachments', False) and self.config.show_has_attachment:
                    attach_icon_path = resource_path("icon2/@.png")
                    attach_img = None
                    if os.path.exists(attach_icon_path):
                        attach_img = self.load_icon_colored(attach_icon_path, size=(14, 14), color=self.colors.get("accent", "#60CDFF"))
                    if attach_img:
                        lbl_attachment = tk.Label(header_frame, image=attach_img, bg=bg_color)
                        lbl_attachment.image = attach_img
                    else:
                        lbl_attachment = tk.Label(
                            header_frame, 
                            text="@", 
                            fg=self.colors.get("accent", "#60CDFF"), 
                            bg=bg_color, 
                            font=(self.config.font_family, self.config.font_size + 1, "bold"),
                        )
                    lbl_attachment.pack(side="right", padx=(4, 2))
                    ToolTip(lbl_attachment, "Has Attachments")


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
                        font=(self.config.font_family, self.config.font_size + 1, "bold"),
                    )
                    lbl_importance.pack(side="right", padx=(0, 2))


                # Flag Indicator (small icon in header corner)
                if email.get('flag_status', 0) != 0:
                    flag_icon_path = resource_path("icon2/flag.png")
                    if os.path.exists(flag_icon_path):
                        flag_img = self.load_icon_colored(flag_icon_path, size=(14, 14), color="#FF8C00")
                        if flag_img:
                            lbl_flag_icon = tk.Label(header_frame, image=flag_img, bg=bg_color)
                            lbl_flag_icon.image = flag_img
                            lbl_flag_icon.pack(side="right", padx=(2, 2))
                            ToolTip(lbl_flag_icon, "Flagged")

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
                        fg=self.colors["fg_primary"], # Or white if badges are always dark? Let's use fg_primary but badges might need contrast.
                        # Badges have colored backgrounds (Orange, Red, Blue). Text should usually be White.
                        # Exception: Light Yellow categories.
                        # Let's keep "white" for badge text unless we have a specific reason.
                        # Actually self.colors["fg_primary"] is Black in Light mode. White text on Orange badge is good. Black text on Orange badge is also okay.
                        # Let's stick to "white" for now as badge backgrounds are dark/saturated. 
                        # Wait, code already handles logic partially.
                        # StartLine 4694 says fg="white". I'll check if I need to change it.
                        # For now I will leave "white" as it contrasts well with the colored badges.
                        bg=badge_bg, 
                        font=(self.config.font_family, self.config.font_size - 2, "bold"),
                        padx=6, pady=2
                    )
                    lbl_badge.pack(side="right", padx=2)
                    
                # Subject
                if self.config.email_show_subject:
                    lbl_subject = tk.Label(
                        card, 
                        text=email['subject'], 
                        fg=self.colors["fg_secondary"], 
                        bg=bg_color, 
                        font=(self.config.font_family, self.config.font_size),
                        anchor="w",
                        justify="left",
                        wraplength=self.config.width - 40 
                    )
                    lbl_subject.pack(fill="x")
                
                # Preview (Body)
                # Create if either Permanent Show OR Hover Show is enabled
                lbl_preview = None
                # Capture current body lines setting for this card
                try: 
                    lines = int(self.config.email_body_lines)
                except: 
                    lines = 2
                    
                if self.config.email_show_body or self.config.show_hover_content:
                    lbl_preview = tk.Text(
                        card, 
                        height=lines,
                        bg=bg_color, 
                        fg=self.colors["fg_dim"], 
                        font=(self.config.font_family, self.config.font_size - 1),
                        bd=0,
                        highlightthickness=0,
                        wrap="word",
                        cursor="arrow"
                    )
                    # Get preview text or fallback - will be populated on hover if needed
                    preview_text = email.get('body', '').strip() 
                    if not preview_text:
                        preview_text = ""
                    else:
                        # Strip empty lines for cleaner display
                        preview_text = "\n".join(line for line in preview_text.splitlines() if line.strip())
                    
                    lbl_preview.insert("1.0", preview_text)
                    lbl_preview.config(state="disabled") # Read-only
                    
                    # Check if we should initially pack it (Show Body = True)
                    if self.config.email_show_body:
                         lbl_preview.pack(fill="x")
                
                # Icon Cache for this refresh cycle
                # We reuse the main cache but ensure lookups are safe
                def get_cached_icon(path, color, size=(24,24)):
                    key = (path, color)
                    # Use self.image_cache (the main one)
                    if key not in self.image_cache:
                        if os.path.exists(path):
                            self.image_cache[key] = self.load_icon_colored(path, size=size, color=color)
                        else:
                            self.image_cache[key] = None
                    return self.image_cache[key]
                
                # --- Action Frame (Buttons) ---
                # Rename locally to frame_buttons to match references
                frame_buttons = tk.Frame(card, bg=bg_color)
                
                # Populate buttons first (so they exist for binding)
                # Filter for valid buttons (Must have Icon AND Action)
                valid_buttons = [
                    conf for conf in self.config.btn_config 
                    if conf.get("icon") and conf.get("action1") != "None"
                ]

                for conf in valid_buttons:
                    icon = conf.get("icon", "ðŸ”˜")
                    
                    is_png = icon.lower().endswith(".png")
                    btn_image = None
                    
                    if is_png:
                        path = resource_path(os.path.join("icons", icon)) 
                        
                        # Let's try to map "white" to a color that works for the theme
                        btn_color = self.colors.get("fg_text", "#FFFFFF")
                        
                        btn_image = get_cached_icon(path, color=btn_color, size=(24, 24))
                    
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
                            fg=self.colors["fg_primary"], 
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
                    btn.bind("<Enter>", lambda e, b=btn: b.config(bg=self.colors["bg_card_hover"]))
                    btn.bind("<Leave>", lambda e, b=btn, bg=bg_color: b.config(bg=bg))
                    
                    # Tooltip logic
                    act1 = conf.get("action1", "")
                    act2 = conf.get("action2", "None")
                    tip_text = "{} & {}".format(act1, act2) if act2 != "None" else act1
                    # Show 'Un-flag' tooltip if email is already flagged
                    if act1 == "Flag" and email.get('flag_status', 0) != 0:
                        tip_text = tip_text.replace('Flag', 'Un-flag')
                    ToolTip(btn, tip_text)
                    
                    # Bind Action (pass card widget for instant removal)
                    btn.bind("<Button-1>", lambda e, c=conf, em=email, w=card: self.handle_custom_action(c, em, source_card=w))

                # --- Logic for Buttons Visibility ---
                if self.config.buttons_on_hover:
                    # Start hidden
                    frame_buttons.pack_forget()
                else:
                    # Always show
                    frame_buttons.pack(fill="x", expand=True, padx=2, pady=(0, 2))


                # --- HOVER BINDINGS (Content & Buttons) ---
                # Define common show/hide helpers with DEFAULT ARGS to capture loop variables correctly
                # We also capture 'lines' from the scope to ensure correct height
                def show_hover_elements(e, lp=lbl_preview, fb=frame_buttons, h=lines, eid=email.get('entry_id'), sid=email.get('store_id')):
                    # 1. Show Body Preview if enabled and not permanent
                    if self.config.show_hover_content and not self.config.email_show_body and lp:
                         # Lazy-fetch body on first hover
                         if not getattr(lp, '_body_loaded', False):
                             lp._body_loaded = True
                             try:
                                 item = self.outlook_client.get_item_by_entryid(eid, sid)
                                 if item:
                                     body_text = ""
                                     try:
                                         body_text = item.Body or ""
                                     except: pass
                                     # Clean up plain text body: remove standalone URLs (tracking links, etc.)
                                     if body_text:
                                         import re
                                         # Remove lines that are just URLs
                                         body_text = re.sub(r'^\s*https?://\S+\s*$', '', body_text, flags=re.MULTILINE)
                                         # Remove inline URLs (but keep surrounding text)
                                         body_text = re.sub(r'https?://\S+', '', body_text)
                                         body_text = body_text.strip()
                                     # If plain body is too short, try extracting from HTML
                                     if len(body_text.strip()) < 30:
                                         try:
                                             import re
                                             html = item.HTMLBody or ""
                                             # Replace <a> tags with their display text (not the href URL)
                                             text = re.sub(r'<a[^>]*>(.*?)</a>', r'\1', html, flags=re.DOTALL|re.IGNORECASE)
                                             # Remove style blocks
                                             text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)
                                             # Remove script blocks
                                             text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL)
                                             # Strip remaining HTML tags
                                             text = re.sub(r'<[^>]+>', ' ', text)
                                             # Decode HTML entities
                                             text = re.sub(r'&nbsp;', ' ', text)
                                             text = re.sub(r'&amp;', '&', text)
                                             text = re.sub(r'&lt;', '<', text)
                                             text = re.sub(r'&gt;', '>', text)
                                             text = re.sub(r'&#\d+;', '', text)
                                             # Remove any remaining URLs
                                             text = re.sub(r'https?://\S+', '', text)
                                             # Collapse whitespace
                                             text = re.sub(r'[ \t]+', ' ', text)
                                             text = re.sub(r'\n\s*\n', '\n', text)
                                             text = text.strip()
                                             if len(text) > len(body_text.strip()):
                                                 body_text = text
                                         except: pass
                                     if body_text:
                                         # Strip empty lines
                                         body_text = "\n".join(line for line in body_text.strip().splitlines() if line.strip())
                                         lp.config(state="normal")
                                         lp.delete("1.0", "end")
                                         lp.insert("1.0", body_text)
                                         lp.config(state="disabled")
                             except Exception as ex:
                                 print("Hover body fetch error: {}".format(ex))
                         # Auto-size: count actual lines of content
                         if not lp.winfo_ismapped():
                              try:
                                  content = lp.get("1.0", "end-1c")
                                  line_count = max(content.count("\n") + 1, 2)
                                  hover_h = min(line_count, 12)  # Cap at 12 lines
                              except:
                                  hover_h = 4
                              lp.config(height=hover_h) 
                              lp.pack(fill="x", padx=5, pady=(0, 2)) 
                    
                    # 2. Show Buttons if enabled
                    if self.config.buttons_on_hover:
                         if not fb.winfo_ismapped():
                              fb.pack(fill="x", expand=True, padx=2, pady=(0, 2))
                
                def hide_hover_elements(e, lp=lbl_preview, fb=frame_buttons):
                    # 1. Hide Body Preview
                    if self.config.show_hover_content and not self.config.email_show_body and lp:
                         if lp.winfo_ismapped():
                              lp.pack_forget()
                    
                    # 2. Hide Buttons
                    if self.config.buttons_on_hover:
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
                
                def safe_show(e, c=card, lp=lbl_preview, fb=frame_buttons, _shf=show_hover_elements):
                     # Delay show to prevent flashing (Debounce)
                     if hasattr(c, "_show_timer") and c._show_timer:
                         c.after_cancel(c._show_timer)
                     c._show_timer = c.after(250, lambda: _shf(e, lp, fb))

                # Apply Bindings
                if (self.config.show_hover_content and not self.config.email_show_body) or self.config.buttons_on_hover:
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
                if self.config.email_double_click: 
                     card.bind("<Double-Button-1>", on_card_click)
                     card.bind("<Button-1>", lambda e, c=card: c.focus_set())
                else:
                     card.bind("<Button-1>", on_card_click)
                
                # Bind Children (Robustly)
                for child in card.winfo_children():
                     # Don't bind click to buttons (they have their own actions)
                     if child != frame_buttons and getattr(child, "master", None) != frame_buttons:
                        if self.config.email_double_click: 
                            child.bind("<Double-Button-1>", on_card_click)
                        else:
                            child.bind("<Button-1>", on_card_click)
                     # Preview text click -> Open Email
                     if child == lbl_preview:
                          if self.config.email_double_click: 
                              child.bind("<Double-Button-1>", on_card_click)
                          else:
                              child.bind("<Button-1>", on_card_click)

                
                if self.config.email_double_click: 
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

            # --- Anti-flicker: reveal rebuilt content in one step ---
            canvas.itemconfigure(self.scroll_frame.window_id, state='normal')
            canvas.configure(scrollregion=canvas.bbox('all'))

            # Ensure Reminders are also refreshed (skip for non-flag email actions)
            if not skip_reminders:
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
        if not self.outlook_client: return
        # Ensure scrollable frame exists
        # Clear content
        if self.reminder_list:
            # --- Anti-flicker: hide canvas content during rebuild ---
            r_canvas = self.reminder_list.canvas
            r_canvas.itemconfigure(self.reminder_list.window_id, state='hidden')
            for widget in self.reminder_list.scrollable_frame.winfo_children():
                widget.destroy()
        
        container = self.reminder_list.scrollable_frame
        
        # Helper for binding click
        def bind_click(widget, entry_id):
            if self.config.email_double_click:
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
        
        if "Today" in self.config.reminder_meeting_dates:
             # Today EOD
             d = today_start + timedelta(days=1) - timedelta(seconds=1)
             if d > end_date: end_date = d
             has_date_filter = True
             
        if "Tomorrow" in self.config.reminder_meeting_dates:
             # Tomorrow EOD
             d = today_start + timedelta(days=2) - timedelta(seconds=1)
             if d > end_date: end_date = d
             has_date_filter = True

        if "Next 7 Days" in self.config.reminder_meeting_dates:
             d = today_start + timedelta(days=8) - timedelta(seconds=1) # Today + 7 full days
             if d > end_date: end_date = d
             has_date_filter = True
             
        if "Custom" in self.config.reminder_meeting_dates:
             try:
                 days = int(getattr(self.config, "reminder_custom_days", 30))
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
             cal_accounts = [n for n, s in self.config.enabled_accounts.items() if s.get("calendar")] if self.config.enabled_accounts else None
             # Pass datetime objects directly
             raw_meetings = self.outlook_client.get_calendar_items(today_start, end_date, cal_accounts)
             
             # Filter by Status
             # olResponseNone = 0, olResponseOrganized = 1, olResponseTentative = 2, olResponseAccepted = 3, olResponseDeclined = 4
             meetings = []
             for m in raw_meetings:
                 status = m.get("response_status", 0)
                 
                 # Accepted
                 if status == 3 and self.config.reminder_accepted_meetings:
                     meetings.append(m)
                     continue
                     
                 # Declined
                 if status == 4 and self.config.reminder_declined_meetings:
                     meetings.append(m)
                     continue
                     
                 # Pending (None, Organized, Tentative, NotResponded=5)
                 # Basically anything not Accepted(3) or Declined(4)
                 if status not in [3, 4] and self.config.reminder_pending_meetings:
                     meetings.append(m)
                     continue
        
        # Icon Cache for this refresh cycle
        icon_cache = {}
        def get_cached_icon(path, color, size=(24,24)):
            key = (path, color, size)
            if key not in icon_cache:
                if os.path.exists(path):
                    icon_cache[key] = self.load_icon_colored(path, size=size, color=color)
                else:
                    icon_cache[key] = None
            return icon_cache[key]

        if meetings:
            tk.Label(container, text="CALENDAR", fg="#60CDFF", bg=self.colors["bg_root"], font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(5, 2))
            # Filter out dismissed items
            meetings = [m for m in meetings if m.get('entry_id') not in self.dismissed_calendar_ids]
            
            # Reset calendar widget tracking for urgency updates
            self._calendar_widgets = []

            for m in meetings:
                 mf = tk.Frame(container, bg=self.colors["bg_card"], padx=5, pady=5)
                 mf.pack(fill="x", padx=2, pady=1)
                 
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
                 
                 # --- Urgency Color Logic ---
                 time_fg, subj_fg = self._get_cal_urgency_colors(m.get('start'))
                     
                 # --- Calendar Buttons Frame (pack RIGHT first so it reserves space) ---
                 c_actions = tk.Frame(mf, bg=self.colors["bg_card"])

                 def make_cal_btn(parent, text, cmd, tip):
                     btn = tk.Label(parent, text=text, fg=self.colors["fg_dim"], bg=self.colors["bg_card"], font=("Segoe UI", 10), cursor="hand2", padx=3)
                     btn.pack(side="right", padx=2)
                     btn.bind("<Button-1>", lambda e: cmd())
                     btn.bind("<Enter>", lambda e: btn.config(fg=self.colors["fg_primary"], bg=self.colors["bg_card_hover"]))
                     btn.bind("<Leave>", lambda e: btn.config(fg=self.colors["fg_dim"], bg=self.colors["bg_card"]))
                     if tip: ToolTip(btn, tip)
                     return btn

                 def do_dismiss_cal(eid=m['entry_id'], w=mf):
                     self.dismissed_calendar_ids.add(eid)
                     w.pack_forget()

                 btn_dismiss = None
                 if os.path.exists(resource_path("icon2/tick-box.png")):
                      img = get_cached_icon(resource_path("icon2/tick-box.png"), color=self.colors["fg_dim"])
                      if img:
                          btn_dismiss = tk.Label(c_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=3)
                          btn_dismiss.image = img

                 if not btn_dismiss:
                      btn_dismiss = make_cal_btn(c_actions, u"âœ“", do_dismiss_cal, "Dismiss")
                 else:
                      btn_dismiss.pack(side="right", padx=2)
                      btn_dismiss.bind("<Button-1>", lambda e, _f=do_dismiss_cal: _f())
                      ToolTip(btn_dismiss, "Dismiss")

                 # Open Meeting Button - use PNG icon
                 btn_open_cal = None
                 if os.path.exists(resource_path("icon2/open-task.png")):
                      img = get_cached_icon(resource_path("icon2/open-task.png"), color=self.colors["fg_dim"], size=(20,20))
                      if img:
                          btn_open_cal = tk.Label(c_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=3)
                          btn_open_cal.image = img
                 
                 if not btn_open_cal:
                      make_cal_btn(c_actions, "Open", lambda eid=m['entry_id'], wlink=m.get('web_link'): self.open_email(eid, fallback_link=wlink), "Open Meeting")
                 else:
                      btn_open_cal.pack(side="right", padx=2)
                      btn_open_cal.bind("<Button-1>", lambda e, eid=m['entry_id'], wlink=m.get('web_link'): self.open_email(eid, fallback_link=wlink))
                      btn_open_cal.bind("<Enter>", lambda e, b=btn_open_cal: b.config(bg=self.colors["bg_card_hover"]))
                      btn_open_cal.bind("<Leave>", lambda e, b=btn_open_cal: b.config(bg=self.colors["bg_card"]))
                      ToolTip(btn_open_cal, "Open Meeting")

                 # Pack buttons frame RIGHT first, before labels
                 c_actions.pack(side="right", padx=2)

                 # Now pack labels LEFT (remaining space after buttons)
                 time_lbl = tk.Label(mf, text=time_str, fg=time_fg, bg=self.colors["bg_card"], font=("Segoe UI", 9))
                 time_lbl.pack(side="left")
                 subj = tk.Label(mf, text=m['subject'], fg=subj_fg, bg=self.colors["bg_card"], font=("Segoe UI", 9, "bold"), anchor="w", wraplength=self.config.width - 130)
                 subj.pack(side="left", padx=5)
                 
                 # Track for live urgency updates
                 self._calendar_widgets.append((m.get('start'), time_lbl, subj, mf))

                 bind_click(mf, m['entry_id'])
                 bind_click(subj, m['entry_id'])

                 # Start hidden, show on hover
                 c_actions.pack_forget()

                 def show_c_actions(e, fa=c_actions):
                     if not fa.winfo_ismapped():
                         fa.pack(side="right", padx=2)

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

                 mf.bind("<Enter>", show_c_actions)
                 mf.bind("<Leave>", hide_c_actions)
                 subj.bind("<Enter>", show_c_actions)
                 subj.bind("<Leave>", hide_c_actions)
            
            # Start the urgency timer if we have calendar widgets
            self._start_cal_urgency_timer()

        # 2. Outlook Tasks
        # 2. Outlook Tasks
        if self.config.reminder_show_tasks:
             tasks = self.outlook_client.get_tasks(due_filters=self.config.reminder_task_dates, account_names=cal_accounts)
             
             if tasks:
                 tk.Label(container, text="TASKS", fg=self.colors.get("accent_success", "#28a745"), bg=self.colors["bg_root"], font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(10, 2))
                 for task in tasks:
                     # Determine if task is overdue
                     t_overdue = False
                     try:
                         if task.get('due') and hasattr(task['due'], 'date'):
                             t_overdue = task['due'].date() < datetime.now().date()
                     except: pass
                     
                     border_color = "#FF6B6B" if t_overdue else self.colors.get("accent_success", "#28a745")
                     tf = tk.Frame(container, bg=self.colors["bg_card"], highlightthickness=1, highlightbackground=border_color, padx=5, pady=5)
                     tf.pack(fill="x", padx=2, pady=2)
                     
                     # Task Buttons Frame (pack RIGHT first so it reserves space)
                     t_actions = tk.Frame(tf, bg=self.colors["bg_card"])


                     # Helper to create buttons
                     def make_task_btn(parent, text, cmd, tip):
                         btn = tk.Label(parent, text=text, fg=self.colors["fg_dim"], bg=self.colors["bg_card"], font=("Segoe UI", 10), cursor="hand2", padx=5)
                         btn.pack(side="right", padx=5) # Align Right
                         btn.bind("<Button-1>", lambda e: cmd())
                         btn.bind("<Enter>", lambda e: btn.config(fg=self.colors["fg_primary"], bg=self.colors["bg_card_hover"]))
                         btn.bind("<Leave>", lambda e: btn.config(fg=self.colors["fg_dim"], bg=self.colors["bg_card"]))
                         if tip: ToolTip(btn, tip)
                         return btn

                     # Complete Button (Checkmark) - Far Right
                     def do_complete(eid=task['entry_id'], sid=task.get('store_id'), w=tf):
                         success = self.outlook_client.mark_task_complete(eid, sid)
                         if success:
                             # Fade out or remove
                             w.pack_forget()
                             # message?
                         else:
                             messagebox.showerror("Error", "Failed to mark task complete.")
                     
                     # Try PNG for complete
                     btn_complete = None
                     if os.path.exists(resource_path("icon2/tick-box.png")):
                          img = get_cached_icon(resource_path("icon2/tick-box.png"), color=self.colors["fg_dim"])
                          if img:
                              btn_complete = tk.Label(t_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=5)
                              btn_complete.image = img # Keep ref
                     
                     if not btn_complete:
                          make_task_btn(t_actions, u"✓", do_complete, "Mark Complete")
                     else:
                          btn_complete.pack(side="right", padx=5)
                          btn_complete.bind("<Button-1>", lambda e, _f=do_complete: _f())
                          ToolTip(btn_complete, "Mark Complete")

                     # Open Button (Folder icon or similar) - Left of Complete
                     btn_open = None
                     if os.path.exists(resource_path("icon2/open-task.png")):
                          img = get_cached_icon(resource_path("icon2/open-task.png"), color=self.colors["fg_dim"], size=(20,20))
                          if img:
                              btn_open = tk.Label(t_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=5)
                              btn_open.image = img
                     
                     if not btn_open:
                          make_task_btn(t_actions, u"📂", lambda eid=task['entry_id'], wlink=task.get('web_link'): self.open_email(eid, fallback_link=wlink), "Open Task")
                     else:
                          btn_open.pack(side="right", padx=5)
                          btn_open.bind("<Button-1>", lambda e, eid=task['entry_id'], wlink=task.get('web_link'): self.open_email(eid, fallback_link=wlink))
                          ToolTip(btn_open, "Open Task")
                     

                     # Pack task buttons frame RIGHT first
                     t_actions.pack(side="right", padx=2)

                     # Task Date
                     try:
                         t_date_str = ""
                         if task.get('due'):
                             t_dt = task['due']
                             t_now = datetime.now()
                             t_is_today = t_dt.date() == t_now.date()
                             t_is_tomorrow = t_dt.date() == (t_now.date() + timedelta(days=1))
                             
                             if t_is_today:
                                 t_date_str = "Today"
                             elif t_is_tomorrow:
                                 t_date_str = "Tomorrow"
                             else:
                                 t_date_str = t_dt.strftime("%a %d")
                         
                         if t_date_str:
                             date_fg = "#FF6B6B" if t_overdue else self.colors["fg_dim"]
                             tk.Label(tf, text=t_date_str, fg=date_fg, bg=self.colors["bg_card"], font=("Segoe UI", 9)).pack(side="left")
                     except: pass

                     subj = tk.Label(tf, text=task['subject'], fg=self.colors["fg_primary"], bg=self.colors["bg_card"], font=("Segoe UI", 9), anchor="w", justify="left", wraplength=self.config.width-130)
                     subj.pack(side="left", padx=5, pady=(0, 2))

                     bind_click(tf, task['entry_id'])
                     bind_click(subj, task['entry_id'])

                     # Start hidden, show on hover
                     t_actions.pack_forget()

                     def show_t_actions(e, fa=t_actions):
                         if not fa.winfo_ismapped():
                             fa.pack(side="right", padx=2)

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
        if self.config.reminder_show_flagged:
             email_accounts = [n for n, s in self.config.enabled_accounts.items() if s.get("email")] if self.config.enabled_accounts else None
             flags, _ = self.outlook_client.get_inbox_items(
                 count=30,
                 unread_only=False,
                 only_flagged=True,
                 due_filters=self.config.reminder_due_filters,
                 account_names=email_accounts
             )
             
             if flags:
                 tk.Label(container, text="FLAGGED EMAILS", fg=self.colors.get("accent_warning", "#FF8C00"), bg=self.colors["bg_root"], font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", padx=5, pady=(10, 2))
                 
                 for email in flags:
                     cf = tk.Frame(container, bg=self.colors["bg_card"], highlightthickness=1, highlightbackground=self.colors.get("accent_warning", "#FF8C00"), padx=5, pady=5)
                     cf.pack(fill="x", padx=2, pady=2)
                     
                     # Subject Label (Now packed top)
                     subj = tk.Label(cf, text=email['subject'], fg=self.colors["fg_primary"], bg=self.colors["bg_card"], font=("Segoe UI", 9), anchor="w", justify="left", wraplength=self.config.width-40)
                     subj.pack(side="top", fill="x", expand=True, padx=5, pady=(0, 2))
                     
                     bind_click(cf, email['entry_id'])
                     bind_click(subj, email['entry_id'])

                     # Tooltip showing flag request and due date
                     tip_parts = []
                     if email.get('flag_request'):
                         tip_parts.append(email['flag_request'])
                     due = email.get('due_date')
                     if due:
                         try:
                             if due.year < 3000:
                                 tip_parts.append("Due: {}".format(due.strftime('%a %d %b %Y')))
                             else:
                                 tip_parts.append("No due date")
                         except:
                             pass
                     if tip_parts:
                         ToolTip(cf, "\n".join(tip_parts))
                         ToolTip(subj, "\n".join(tip_parts))

                     # Flag Actions Frame (Hidden initially)
                     # Packed below subject
                     f_actions = tk.Frame(cf, bg=self.colors["bg_card"])
                     # f_actions.pack(side="top", fill="x", padx=2) # Hide by default

                     # Helper to create buttons
                     def make_flag_btn(parent, text, cmd, tip):
                         btn = tk.Label(parent, text=text, fg=self.colors["fg_dim"], bg=self.colors["bg_card"], font=("Segoe UI", 10), cursor="hand2", padx=5)
                         btn.pack(side="right", padx=5) # Pack right for alignment
                         btn.bind("<Button-1>", lambda e, _c=cmd: _c())
                         btn.bind("<Enter>", lambda e: btn.config(fg=self.colors["fg_primary"], bg=self.colors["bg_card_hover"]))
                         btn.bind("<Leave>", lambda e: btn.config(fg=self.colors["fg_dim"], bg=self.colors["bg_card"]))
                         if tip: ToolTip(btn, tip)
                         return btn

                     # Unflag Button (Flag icon) - Moved to far right (first packed right)
                     def do_unflag(eid=email['entry_id'], sid=email['store_id'], w=cf):
                         success = self.outlook_client.unflag_email(eid, sid)
                         if success:
                             w.pack_forget()
                             self.refresh_emails()  # Update email list to remove flag indicator
                         else:
                             messagebox.showerror("Error", "Failed to unflag email.")
                     
                     btn_unflag = None
                     if os.path.exists(resource_path("icon2/flag.png")):
                          img = get_cached_icon(resource_path("icon2/flag.png"), color="#FF8C00")
                          if img:
                              btn_unflag = tk.Label(f_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=5)
                              btn_unflag.image = img
                     
                     if not btn_unflag:
                          make_flag_btn(f_actions, u"⚐", do_unflag, "Unflag")
                     else:
                          # Pack Right
                          btn_unflag.pack(side="right", padx=5)
                          btn_unflag.bind("<Button-1>", lambda e, _f=do_unflag: _f())
                          ToolTip(btn_unflag, "Unflag")

                     # Open Button (Folder icon)
                     btn_open = None
                     if os.path.exists(resource_path("icon2/open-task.png")):
                          img = get_cached_icon(resource_path("icon2/open-task.png"), color=self.colors["fg_dim"], size=(20,20))
                          if img:
                              btn_open = tk.Label(f_actions, image=img, bg=self.colors["bg_card"], cursor="hand2", padx=5)
                              btn_open.image = img
                     
                     if not btn_open:
                          make_flag_btn(f_actions, u"📂", lambda eid=email['entry_id']: self.open_email(eid), "Open Email")
                     else:
                          # Pack Right (left/next to Unflag)
                          btn_open.pack(side="right", padx=5)
                          btn_open.bind("<Button-1>", lambda e, eid=email['entry_id']: self.open_email(eid))
                          ToolTip(btn_open, "Open Email")

                     # Flag info labels (packed LEFT, so they appear to the left of buttons)
                     flag_req = email.get('flag_request', '')
                     # Abbreviate common flag requests
                     abbrev = {
                         'For Your Information': 'FYI',
                         'Follow up': 'Follow up',
                         'Forward': 'Forward',
                         'Review': 'Review',
                         'Reply': 'Reply',
                         'Reply to All': 'Reply All',
                         'Call': 'Call',
                         'Do not Forward': 'No Forward',
                         'Read': 'Read',
                     }
                     short_req = abbrev.get(flag_req, flag_req) if flag_req else ''

                     # Build due date text
                     due_text = ''
                     due = email.get('due_date')
                     if due:
                         try:
                             # Strip timezone info if present (COM returns tz-aware datetimes)
                             if hasattr(due, 'tzinfo') and due.tzinfo:
                                 due = due.replace(tzinfo=None)
                             if due.year < 3000:
                                 now_dt = datetime.now()
                                 today_d = now_dt.replace(hour=0, minute=0, second=0, microsecond=0)
                                 due_d = due.replace(hour=0, minute=0, second=0, microsecond=0)
                                 diff = (due_d - today_d).days
                                 if diff < 0:
                                     due_text = "Overdue {}d".format(abs(diff))
                                 elif diff == 0:
                                     due_text = "Today"
                                 elif diff == 1:
                                     due_text = "Tomorrow"
                                 else:
                                     due_text = due.strftime('%a %d %b')
                         except:
                             pass

                     info_text = ''
                     if short_req and due_text:
                         info_text = "{} · {}".format(short_req, due_text)
                     elif short_req:
                         info_text = short_req
                     elif due_text:
                         info_text = due_text

                     if info_text:
                         due_color = "#FF6B6B" if due_text.startswith("Overdue") else self.colors["fg_dim"]
                         lbl_info = tk.Label(f_actions, text=info_text, fg=due_color, bg=self.colors["bg_card"], font=("Segoe UI", 8), anchor="w")
                         lbl_info.pack(side="left", padx=(5, 0))

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

        # --- Anti-flicker: reveal rebuilt reminders content in one step ---
        if self.reminder_list:
            r_canvas = self.reminder_list.canvas
            r_canvas.itemconfigure(self.reminder_list.window_id, state='normal')
            r_canvas.configure(scrollregion=r_canvas.bbox('all'))


    def draw_pin_icon(self):
        # Determine active icon based on state
        if isinstance(self.btn_pin, tk.Label):
             # If using Label with Images
             if self.config.pinned:
                 if hasattr(self, 'icon_pin_active'):
                     self.btn_pin.config(image=self.icon_pin_active)
             else:
                 if hasattr(self, 'icon_pin_inactive'):
                     self.btn_pin.config(image=self.icon_pin_inactive)
        
        elif isinstance(self.btn_pin, tk.Canvas):
            # If using Canvas drawing
            self.btn_pin.delete("all")
            color = "#007ACC" if self.config.pinned else "#AAAAAA"
            # Draw a simple pin shape
            self.btn_pin.create_oval(10, 5, 20, 15, fill=color, outline="")
            self.btn_pin.create_line(15, 15, 15, 25, fill=color, width=2)

    def toggle_pin(self):
        self.config.pinned = not self.config.pinned
        self.config.save()
        
        if self.toolbar:
            self.toolbar.update_pin_state()
            
        self.apply_state()
        
        # Force a check immediately to update pulse state
        self.after(100, self._perform_check)

    def apply_state(self):
        """Applies the current state (Pinned/Expanded/Collapsed) to the window and AppBar."""
        # Update monitor info to ensure correct sizing on monitor change
        self.monitor_x, self.monitor_y, self.screen_width, self.screen_height = self.get_current_monitor_info()

        # Update AppBar edge based on preference
        new_edge = ABE_LEFT if self.config.dock_side == "Left" else ABE_RIGHT
        
        # If side changed, we MUST unregister the old one first to release the old edge
        if self.appbar.edge != new_edge:
            self.appbar.unregister()
            self.appbar.edge = new_edge
            self.appbar.abd.uEdge = new_edge

        if self.config.pinned:
            # Pinned: Reserve Full Width, Visual Full Width
            mode = "PINNED"
            reserve_w = self.config.width
            visual_w = self.config.width
            
        elif self.is_expanded:
            # Expanded/Overlay: Reserve Strip Width, Visual Full Width
            # This allows the "Overlay" effect without losing the "Sneak Behind" protection for the strip.
            mode = "OVERLAY"
            reserve_w = self.hot_strip_width
            visual_w = self.config.width
            
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
            if self.config.dock_side == "Left":
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
        if self.config.pinned or self.is_expanded:
            x_root = self.winfo_pointerx()
            
            # Calculate width based on side
            if self.config.dock_side == "Left":
                new_width = x_root - self.monitor_x
            else:
                new_width = (self.monitor_x + self.screen_width) - x_root
            
            if new_width > self.min_width and new_width < (self.screen_width // 2):
                self.config.width = new_width
                # Optimization: ONLY resize the visual window, do NOT trigger AppBar reflow
                self.set_geometry(self.config.width)
                # Ensure the content knows we resized if needed (pack handles this)

    def on_resize_release(self, event):
        # Commit the new width to the system (triggers reflow once)
        self.apply_state() 
        self.config.save()

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
            
            if self.config.dock_side == "Right":
                 x = rect.right - width
        else:
            # Unpinned / Overlay Mode (Only if registration failed or deliberately unregistered)
            # Use Work Area for Height/Y to respect Taskbar
            y = wy
            h = wh
            
            # X Calculation: Anchor to Monitor Edge
            # But ensure we don't start 'under' a vertical taskbar on the left?
            # Using Work Area left (wx) is safest.
            
            if self.config.dock_side == "Left":
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
        Uses MonitorFromWindow to find the monitor containing the window.
        Returns check-safe dictionary with 'monitor' and 'work' tuples (x,y,w,h).
        """
        try:
            user32 = ctypes.windll.user32
            if hasattr(self, 'hwnd') and self.hwnd:
                hwnd = self.hwnd
            else:
                hwnd = self.winfo_id()
            
            # MONITOR_DEFAULTTONEAREST = 2
            hMonitor = user32.MonitorFromWindow(hwnd, 2)
            
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            
            if user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi)):
                r = mi.rcMonitor
                w = mi.rcWork
                return {
                    "monitor": (r.left, r.top, r.right - r.left, r.bottom - r.top),
                    "work":    (w.left, w.top, w.right - w.left, w.bottom - w.top),
                }
        except Exception as e:
            print("Error in get_monitor_metrics: {}".format(e))

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
            self.config.dock_side = "Left"
        else:
            self.config.dock_side = "Right"
            
        # Re-apply state which will snap to monitor edge and re-register
        self.apply_state()

    # --- Calendar Urgency ---
    def _get_cal_urgency_colors(self, start_dt):
        """Returns (time_fg, subj_fg) colors based on meeting proximity.
        
        - Default: dim time, normal subject
        - 15 min before: amber time, amber subject
        - At meeting time (0 to +1 min): orange/pulsing
        - 1+ min overdue: red
        """
        if not start_dt or not hasattr(start_dt, 'timestamp'):
            return self.colors["fg_dim"], self.colors["fg_primary"]
        
        try:
            from datetime import datetime
            now = datetime.now()
            # Strip timezone info if present (Outlook returns tz-aware datetimes)
            meeting_time = start_dt.replace(tzinfo=None) if hasattr(start_dt, 'tzinfo') and start_dt.tzinfo else start_dt
            diff_seconds = (meeting_time - now).total_seconds()
            diff_minutes = diff_seconds / 60.0
            
            if diff_minutes <= -1:
                # 1+ minute overdue → red
                return "#FF4444", "#FF4444"
            elif diff_minutes <= 0:
                # At meeting time (0 to -1 min) → bright orange
                return "#FF8C00", "#FF8C00"
            elif diff_minutes <= 15:
                # Within 15 minutes → amber
                return "#FFB347", "#FFB347"
            else:
                # Normal
                return self.colors["fg_dim"], self.colors["fg_primary"]
        except:
            return self.colors["fg_dim"], self.colors["fg_primary"]
    
    def _update_cal_urgency(self):
        """Update all tracked calendar widgets with current urgency colors."""
        for entry in self._calendar_widgets:
            try:
                start_dt, time_lbl, subj_lbl, frame = entry
                # Check widget still exists
                if not time_lbl.winfo_exists():
                    continue
                time_fg, subj_fg = self._get_cal_urgency_colors(start_dt)
                time_lbl.config(fg=time_fg)
                subj_lbl.config(fg=subj_fg)
            except:
                pass
    
    def _start_cal_urgency_timer(self):
        """Start periodic urgency color updates (every 30 seconds)."""
        if self._cal_urgency_timer:
            try:
                self.after_cancel(self._cal_urgency_timer)
            except:
                pass
        
        self._cal_pulse_on = True
        
        def tick():
            if not self._calendar_widgets:
                return
            try:
                from datetime import datetime
                now = datetime.now()
                
                for entry in self._calendar_widgets:
                    try:
                        start_dt, time_lbl, subj_lbl, frame = entry
                        if not time_lbl.winfo_exists():
                            continue
                        
                        meeting_time = start_dt.replace(tzinfo=None) if hasattr(start_dt, 'tzinfo') and start_dt.tzinfo else start_dt
                        diff_seconds = (meeting_time - now).total_seconds()
                        diff_minutes = diff_seconds / 60.0
                        
                        if -1 <= diff_minutes <= 0:
                            # Meeting is NOW — pulse between orange and dim
                            if self._cal_pulse_on:
                                time_lbl.config(fg="#FF8C00")
                                subj_lbl.config(fg="#FF8C00")
                            else:
                                time_lbl.config(fg=self.colors["fg_dim"])
                                subj_lbl.config(fg=self.colors["fg_dim"])
                        else:
                            # Normal urgency update (no pulse needed)
                            time_fg, subj_fg = self._get_cal_urgency_colors(start_dt)
                            time_lbl.config(fg=time_fg)
                            subj_lbl.config(fg=subj_fg)
                    except:
                        pass
                
                self._cal_pulse_on = not self._cal_pulse_on
            except:
                pass
            
            # Run every 2 seconds for smooth pulsing
            self._cal_urgency_timer = self.after(2000, tick)
        
        tick()

    # --- Polling Control ---
    # --- Polling Control ---
    def start_polling(self):
        """Starts the background polling loop."""
        self.check_updates()
        self.check_fullscreen_app()

    def _check_for_app_update(self):
        """Check GitHub for a newer version (runs in background thread)."""
        def _on_result(latest_version, download_url):
            if latest_version:
                # Schedule UI update on main thread
                self.after(0, lambda: self._show_update_bar(latest_version, download_url))
        
        check_for_update(_on_result)
    
    def _show_update_bar(self, version, download_url):
        """Shows a subtle update notification bar above the footer."""
        if hasattr(self, '_update_bar') and self._update_bar:
            try: self._update_bar.destroy()
            except: pass
        
        bar = tk.Frame(self.main_frame, bg=self.colors["accent"], cursor="hand2")
        bar.pack(fill="x", side="bottom", before=self.footer)
        
        lbl = tk.Label(
            bar,
            text="\u2b06  Update {} available — click to download".format(version),
            bg=self.colors["accent"],
            fg="#000000",
            font=(self.config.font_family, 8, "bold"),
            cursor="hand2",
            pady=4,
        )
        lbl.pack(fill="x")
        
        import webbrowser
        bar.bind("<Button-1>", lambda e: webbrowser.open(download_url))
        lbl.bind("<Button-1>", lambda e: webbrowser.open(download_url))
        
        # Close button
        btn_x = tk.Label(
            bar, text="\u2715", bg=self.colors["accent"], fg="#000000",
            font=(self.config.font_family, 8), cursor="hand2", padx=5,
        )
        btn_x.place(relx=1.0, rely=0.5, anchor="e", x=-2)
        btn_x.bind("<Button-1>", lambda e: bar.destroy())
        
        self._update_bar = bar

    def check_fullscreen_app(self):
        """Checks if a full-screen application is active on the current monitor and helps sidebar get out of the way."""
        # DISABLED: Causes jumpy behavior when switching between maxed apps on different displays
        return
 
    def check_updates(self):
        """Threaded (or scheduled) update check."""
        try:
            self._perform_check()
        except Exception as e:
            print("Polling error: {}".format(e))
        
        # Schedule next poll
        interval = getattr(self.config, "poll_interval", 15) * 1000
        self.after(interval, self.check_updates)

    def _perform_check(self):
        """Actual check logic."""
        if not self.outlook_client: return
        accounts = None
        if self.config.enabled_accounts:
            accounts = list(self.config.enabled_accounts.keys())

        # Safety net: Force a full refresh every 5 minutes regardless
        # This ensures emails recover even if check_new_mail fails silently
        if not hasattr(self, '_last_forced_refresh'):
            self._last_forced_refresh = time.time()
        
        if time.time() - self._last_forced_refresh > 300:
            self._last_forced_refresh = time.time()
            print("DEBUG: Forced periodic refresh (5 min safety net)")
            self.refresh_emails()
            return

        # 1. Check New Mail (For Refreshing List)
        has_new = self.outlook_client.check_new_mail(accounts)
        if has_new:
             print("DEBUG: Refreshing emails...")
             self.refresh_emails()
        
        # 2. Gather Statuses for Pulse
        unread_count = self.outlook_client.get_unread_count(accounts, self.config.enabled_accounts)
        due_status = self.outlook_client.get_pulse_status(accounts)

        active_colors = []
        
        # Priority 1: Unread Mail (Blue)
        if unread_count > 0:
            active_colors.append("#0078D4")

        # Priority 2: Meetings (Orange)
        if due_status["calendar"] == "Today" and self.config.reminder_show_meetings and "Today" in self.config.reminder_meeting_dates:
            active_colors.append("#E68D49") # Soft Orange
            
        # Priority 3: Tasks (Green)
        if self.config.reminder_show_tasks:
            t_status = due_status["tasks"]
            if t_status == "Overdue" and "Overdue" in self.config.reminder_due_filters:
                active_colors.append("#28C745")
            elif t_status == "Today" and "Today" in self.config.reminder_due_filters:
                active_colors.append("#28C745")
            
        # Trigger Pulse if needed
        if active_colors and not self.config.pinned and not self.is_expanded:
            # print("DEBUG: Active Pulse Colors: {}".format(active_colors))
            self.start_pulse(active_colors)
        elif not active_colors:
            self.stop_pulse()

    def start_pulse(self, colors):
        """Starts the hot strip pulsing animation with a list of colors."""
        if self.config.pinned or self.is_expanded: return 
        
        # print("DEBUG: start_pulse triggered. Colors={}".format(colors)) 
        
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
        
        if not self.config.pinned and not self.is_expanded:
            # Start hover timer (0.75s delay)
            if not self._hover_timer:
                self._hover_timer = self.after(750, self.do_expand)

    def do_expand(self):
        """Actually expands the sidebar after delay."""
        self._hover_timer = None
        if not self.config.pinned and not self.is_expanded:
            self.stop_pulse() # Stop pulse only when genuinely opening
            self.is_expanded = True
            self.apply_state() # Expand and reserve space

    def on_leave(self, event):
        # Leaving the window to the desktop should collapse.
        # Verify mouse is truly outside the window (not just crossing a child widget boundary).
        if not self.config.pinned:
            # Cancel potential expand timer if we left quickly
            if self._hover_timer:
                self.after_cancel(self._hover_timer)
                self._hover_timer = None
                
            if self.is_expanded:
                # Check if mouse is actually outside the window
                try:
                    x, y = self.winfo_pointerxy()
                    wx = self.winfo_rootx()
                    wy = self.winfo_rooty()
                    ww = self.winfo_width()
                    wh = self.winfo_height()
                    
                    if x < wx or x >= wx + ww or y < wy or y >= wy + wh:
                        # Mouse is truly outside — start collapse timer
                        if self._collapse_timer:
                            self.after_cancel(self._collapse_timer)
                        self._collapse_timer = self.after(self.config.hover_delay, self.do_collapse)
                except:
                    # Fallback: just start the timer
                    if self._collapse_timer:
                        self.after_cancel(self._collapse_timer)
                    self._collapse_timer = self.after(self.config.hover_delay, self.do_collapse)

    def on_motion(self, event):
        # Reset collapse timer only if mouse is inside the window
        if self._collapse_timer:
            try:
                x, y = self.winfo_pointerxy()
                wx = self.winfo_rootx()
                wy = self.winfo_rooty()
                ww = self.winfo_width()
                wh = self.winfo_height()
                
                if wx <= x < wx + ww and wy <= y < wy + wh:
                    self.after_cancel(self._collapse_timer)
                    self._collapse_timer = None
            except:
                pass

    def do_collapse(self):
        if not self.config.pinned:
            # Double-check mouse is still outside before collapsing
            try:
                x, y = self.winfo_pointerxy()
                wx = self.winfo_rootx()
                wy = self.winfo_rooty()
                ww = self.winfo_width()
                wh = self.winfo_height()
                
                if wx <= x < wx + ww and wy <= y < wy + wh:
                    # Mouse came back — don't collapse
                    self._collapse_timer = None
                    return
            except:
                pass
            
            self._collapse_timer = None
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
        """Legacy wrapper: now defers to self.config.load()"""
        self.config.load()

    def save_config(self):
        """Legacy wrapper: now defers to self.config.save()"""
        self.config.save()


    def handle_quick_create(self):
        """Handles the quick create button click."""
        actions = self.config.quick_create_actions
        
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
        # Delegated to Toolbar
        if self.toolbar:
            self.toolbar.update_quick_create_icon(self.colors)
 
    def _execute_quick_action(self, action):
        try:
            self._allow_foreground_for_outlook()
            if action == "New Email":
                self.outlook_client.create_email()
            elif action == "New Meeting":
                self.outlook_client.create_meeting()
            elif action == "New Appointment":
                # Graph doesn't distinguish deeply; map to meeting
                self.outlook_client.create_meeting()
            elif action == "New Task":
                self.outlook_client.create_task()
        except Exception as e:
            print("Quick create error: {}".format(e))
 

    def apply_theme(self):
        """Applies the current theme colors to all UI components."""
        # Clear image cache to force reload with new colors/context
        self.image_cache = {}
        
        c = self.colors
        
        # 1. Main Window & Frames
        self.configure(bg=c["bg_root"])
        self.content_wrapper.config(bg=c["bg_root"])
        self.main_frame.config(bg=c["bg_root"])
        self.footer.config(bg=c["bg_header"])
        self.header.config(bg=c["bg_header"])
        try: self.resize_grip.config(bg=c["fg_dim"])
        except: pass
        
        # 2. Panes & Scroll Frames
        try:
             self.paned_window.config(bg=c["bg_root"])
             self.pane_emails.config(bg=c["bg_root"])
             self.pane_reminders.config(bg=c["bg_root"])
             self.email_list_frame.config(bg=c["bg_root"])
             self.scroll_frame.config(bg=c["bg_root"], sb_bg=c["fg_dim"], sb_trough=c["scroll_bg"])
             self.reminder_list.config(bg=c["bg_root"], sb_bg=c["fg_dim"], sb_trough=c["scroll_bg"])
        except: pass
 
        # 3. Section Headers
        try:
            self.email_header.config(bg=c["bg_card"])
            self.lbl_email_header.config(bg=c["bg_card"], fg=c["fg_text"])
            self.r_header.config(bg=c["bg_card"])
            for child in self.r_header.winfo_children():
                try: child.config(bg=c["bg_card"], fg=c["fg_dim"])
                except: pass
        except: pass
        
        # 4. Header Elements
        try:
            self.lbl_title.config(bg=c["bg_header"], fg=c["fg_primary"])
            self.btn_account_toggle.config(bg=c["bg_card"])
            self._update_arrow_icons()
            
            if self.toolbar:
                self.toolbar.apply_theme(c)

        except Exception as e:
            print("Error updating header/toolbar: {}".format(e))
        
        # 5. Hot strip & canvas
        try: self.hot_strip.config(bg=c["accent"])
        except: pass
        try: self.hot_strip_canvas.config(bg=c["bg_root"])
        except: pass
 
        # 6. Settings Panel (if open, close and re-open with new colors)
        if hasattr(self, "settings_panel") and self.settings_panel and self.settings_panel.winfo_exists():
             try:
                 self.settings_panel.destroy()
                 self.settings_panel = None
                 self.settings_panel_open = False
                 self.toggle_settings_panel()
             except: pass
             
        # 7. Recolor existing email cards in-place (no COM re-fetch)
        try:
            for card in self.scroll_frame.scrollable_frame.winfo_children():
                self._recolor_widget_tree(card, c)
        except: pass
        
        # 8. Recolor existing reminder items in-place
        try:
            for item in self.reminder_list.scrollable_frame.winfo_children():
                self._recolor_widget_tree(item, c)
        except: pass
        
        self.update_idletasks()
    
    def _recolor_widget_tree(self, widget, c):
        """Recursively recolor a widget and all its children for theme change."""
        try:
            wtype = widget.winfo_class()
            if wtype == "Frame":
                try:
                    hb = widget.cget("highlightbackground")
                    # Card with border - update bg and keep accent border for unread
                    widget.config(bg=c["bg_card"])
                    if hb not in (c["accent"], c["card_border"]):
                        widget.config(highlightbackground=c["card_border"])
                except:
                    widget.config(bg=c["bg_root"])
            elif wtype == "Label":
                parent_bg = c["bg_card"]
                try: parent_bg = widget.master.cget("bg")
                except: pass
                widget.config(bg=parent_bg)
                fg = widget.cget("fg")
                if fg in ("#FFFFFF", "#ffffff", "white", "#000000", "#000", "black"):
                    widget.config(fg=c["fg_text"])
                elif fg in ("#CCCCCC", "#cccccc", "#333333"):
                    widget.config(fg=c["fg_secondary"])
                elif fg in ("#999999", "#666666"):
                    widget.config(fg=c["fg_dim"])
            elif wtype == "Text":
                parent_bg = c["bg_card"]
                try: parent_bg = widget.master.cget("bg")
                except: pass
                widget.config(bg=parent_bg, fg=c["fg_dim"])
            elif wtype == "Canvas":
                widget.config(bg=c["bg_card"])
        except: pass
        
        for child in widget.winfo_children():
            self._recolor_widget_tree(child, c)
 
    def toggle_theme(self):
        """Switches between Light and Dark themes."""
        if self.current_theme == "Dark":
            self.current_theme = "Light"
        else:
            self.current_theme = "Dark"
            
        self.colors = self.palettes[self.current_theme]
        self.config.theme = self.current_theme
        self.save_config()
        
        # Apply changes immediately
        self.apply_theme()
        
        # Full refresh to rebuild cards with correct icon colors and Text widget backgrounds
        try:
            self.refresh_emails()
        except: pass
 
 # --- Single Instance Logic (Mutex) ---
 
class SingleInstance:
    """
    Limits application to a single instance using a Named Mutex.
    Safe for MSIX and standard execution.
    """
    def __init__(self, name="Global\\OutlookSidebar_Mutex_v2"):
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
        print("Starting Sidebar...")
        # Check Single Instance
        app_instance = SingleInstance()
        is_running = app_instance.already_running()
        print("Already running: {}".format(is_running))
        if is_running:
            print("Exiting because already running.")
            sys.exit(0)

        # Keep the mutex handle alive for the duration of the app
        print("Launching SidebarWindow...")
        app = SidebarWindow()
        print("Entering mainloop...")
        app.mainloop()

    except Exception as e:
        import traceback
        with open("startup_error.log", "w") as f:
            f.write("Startup Error: {}\n".format(e))
            f.write(traceback.format_exc())

