import tkinter as tk
import os
import sys

# Try to import ToolTip from existing project structure
try:
    from sidebar.ui.widgets.base import ToolTip
except ImportError:
    # Fallback if structure is different
    class ToolTip:
        def __init__(self, widget, text):
            pass

class SidebarToolbar:
    """
    Manages the toolbar buttons (Header and Footer) for the Sidebar application.
    Separates UI creation and event handling from the main window logic.
    """
    def __init__(self, header_frame, footer_frame, callbacks, image_loader, resource_path_func, config_manager):
        """
        Args:
            header_frame (tk.Frame): The frame to place header buttons.
            footer_frame (tk.Frame): The frame to place footer buttons.
            callbacks (dict): Map of command names to functions (e.g. {'refresh': self.refresh_emails}).
            image_loader (func): Function to load colored icons (likely app.load_icon_colored).
            resource_path_func (func): Function to get absolute resource path.
            config_manager (ConfigManager): The configuration manager instance.
        """
        self.header = header_frame
        self.footer = footer_frame
        self.callbacks = callbacks
        self.load_icon = image_loader
        self.resource_path = resource_path_func
        self.config = config_manager
        
        # UI Elements
        self.btn_pin = None
        self.pin_tooltip = None
        self.btn_settings = None
        self.btn_help = None
        self.btn_refresh = None
        self.btn_share = None
        
        self.btn_outlook = None
        self.btn_calendar = None
        self.btn_quick_create = None
        
        self.lbl_version = None
        self.btn_close = None
        
        # Icon Cache (Local to toolbar, or shared?)
        # Better to let the main app handle caching if the loader does it? 
        # Actually standard practice is to keep references here too to prevent GC.
        self.icons = {}
        self.colors = {}

    def create_header_buttons(self, colors):
        """Creates the buttons in the header frame."""
        self.colors = colors
        # 1. Pin Button
        self._create_pin_button(colors)
        
        # 2. Settings Button
        self.btn_settings = self._create_icon_button(
            self.header, "icon2/spanner.png", "settings", colors, 
            self.callbacks.get("settings"), "Settings", size=(22, 22), fallback_text="⚙"
        )
        
        # 3. Help Button
        self.btn_help = tk.Label(self.header, text="?", bg=colors["bg_header"], fg=colors["fg_dim"], font=("Segoe UI", 14, "bold"), cursor="hand2")
        self.btn_help.pack(side="right", padx=(5, 5), pady=5)
        self.btn_help.bind("<Button-1>", lambda e: self.callbacks.get("help")())
        ToolTip(self.btn_help, "Instructions")
        
        # 4. Refresh Button
        self.btn_refresh = self._create_icon_button(
            self.header, "icon2/refresh.png", "refresh", colors,
            self.callbacks.get("refresh"), "Refresh Email List", size=(22, 22), fallback_text="↻"
        )
        
        # 5. Share Button
        self.btn_share = self._create_icon_button(
            self.header, "icon2/share.png", "share", colors,
            self.callbacks.get("share"), "Share InboxBar", size=(20, 20), fallback_text="🔗"
        )

    def create_footer_buttons(self, colors, version_text="v1.0"):
        """Creates the buttons in the footer frame."""
        # Footer is typically: [Version] [Close] (Right) ... [Outlook][Calendar][QuickCreate] (Left/Center?)
        # Let's check original layout.
        # Original: 
        # lbl_version.pack(side="left", padx=10)
        # btn_close.pack(side="right", padx=10)
        # btn_quick_create.pack(side="right", padx=10)
        # btn_calendar.pack(side="right", padx=10)
        # btn_outlook.pack(side="right", padx=10)

        self.lbl_version = tk.Label(self.footer, text=version_text, bg=colors["bg_header"], fg=colors["fg_dim"], font=("Segoe UI", 8))
        self.lbl_version.pack(side="left", padx=10)
        # Close / Exit Button (uses same red X icon as settings panel)
        close_icon_path = self.resource_path("icon2/close-window.png")
        if os.path.exists(close_icon_path):
            try:
                self.icons["close"] = self.load_icon(close_icon_path, size=(28, 28), color="#FF4444")
                self.btn_close = tk.Label(self.footer, image=self.icons["close"], bg=colors["bg_header"], cursor="hand2")
            except:
                self.btn_close = tk.Label(self.footer, text="\u2715", bg=colors["bg_header"], fg="#FF4444", cursor="hand2", font=("Segoe UI", 12))
        else:
            self.btn_close = tk.Label(self.footer, text="\u2715", bg=colors["bg_header"], fg="#FF4444", cursor="hand2", font=("Segoe UI", 12))
        self.btn_close.pack(side="right", padx=10)
        self.btn_close.bind("<Button-1>", lambda e: self.callbacks.get("close")())
        ToolTip(self.btn_close, "Exit InboxBar")
        
        # Quick Create
        self.btn_quick_create = tk.Label(self.footer, bg=colors["bg_header"], cursor="hand2")
        self.btn_quick_create.pack(side="right", padx=10)
        self.btn_quick_create.bind("<Button-1>", lambda e: self.callbacks.get("quick_create")())
        self.update_quick_create_icon(colors)
        ToolTip(self.btn_quick_create, "Quick Create")

        # Calendar
        self.btn_calendar = self._create_icon_button(
            self.footer, "icon2/calendar.png", "calendar", colors,
            self.callbacks.get("calendar"), "Open Calendar", size=(28, 28)
        )
        
        # Outlook
        self.btn_outlook = self._create_icon_button(
            self.footer, "icon2/email.png", "outlook", colors,
            self.callbacks.get("outlook"), "Open Outlook", size=(32, 32)
        )

    def _flash_button(self, btn, flash_color="#555555", duration=150):
        """Briefly flash a button's background to indicate it was clicked."""
        try:
            orig_bg = btn.cget("bg")
            btn.config(bg=flash_color)
            btn.after(duration, lambda: btn.config(bg=orig_bg))
        except Exception:
            pass

    def _create_icon_button(self, parent, icon_path, cache_key, colors, command, current_tooltip, size=(24,24), fallback_text="?"):
        """Helper to create a standard icon button."""
        btn = None
        path = self.resource_path(icon_path)
        
        if os.path.exists(path):
            try:
                img = self.load_icon(path, size=size, color=colors["fg_primary"])
                self.icons[cache_key] = img
                btn = tk.Label(parent, image=img, bg=colors["bg_header"], cursor="hand2")
            except Exception as e:
                print(f"Error loading {icon_path}: {e}")
        
        if not btn:
             # Fallback
             btn = tk.Label(parent, text=fallback_text, bg=colors["bg_header"], fg=colors["fg_dim"], cursor="hand2")
        
        btn.pack(side="right", padx=5)
        if command:
            def on_click(e, b=btn, cmd=command):
                self._flash_button(b)
                cmd()
            btn.bind("<Button-1>", on_click)
        if current_tooltip:
            ToolTip(btn, current_tooltip)
            
        return btn

    def _create_pin_button(self, colors):
        """Creates or recreates the pin button logic."""
        # This is more complex because it changes state (icon and tooltip)
        path = self.resource_path("icon2/pin1.png")
        
        if os.path.exists(path):
             try:
                  self.icon_pin_active = self.load_icon(path, size=(24, 24), color=colors["accent"])
                  self.icon_pin_inactive = self.load_icon(path, size=(24, 24), color=colors["fg_dim"])
                  
                  img = self.icon_pin_active if self.config.pinned else self.icon_pin_inactive
                  self.btn_pin = tk.Label(self.header, image=img, bg=colors["bg_header"], cursor="hand2")
             except Exception as e:
                  print(f"Error loading Pin icon: {e}")
                  # Fallback
                  self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg=colors["bg_header"], highlightthickness=0)
                  self._draw_pin_canvas(colors)
        else:
             self.btn_pin = tk.Canvas(self.header, width=30, height=30, bg=colors["bg_header"], highlightthickness=0)
             self._draw_pin_canvas(colors)
             
        self.btn_pin.pack(side="right", padx=5, pady=5)
        self.btn_pin.bind("<Button-1>", lambda e: self.callbacks.get("toggle_pin")())
        
        initial_tip = "Unpin Window (Current: Pinned)" if self.config.pinned else "Pin Window (Current: Auto-Collapse)"
        self.pin_tooltip = ToolTip(self.btn_pin, initial_tip)

    def _draw_pin_canvas(self, colors):
        """Fallback drawing for pin if image fails."""
        if not isinstance(self.btn_pin, tk.Canvas): return
        
        self.btn_pin.delete("all")
        color = colors.get("accent", "#007ACC") if self.config.pinned else colors.get("fg_dim", "#AAAAAA")
        # Draw a simple pin shape
        self.btn_pin.create_oval(10, 5, 20, 15, fill=color, outline="")
        self.btn_pin.create_line(15, 15, 15, 25, fill=color, width=2)

    def update_pin_state(self):
        """Updates the visual state of the pin button based on config."""
        if isinstance(self.btn_pin, tk.Label):
             if self.config.pinned:
                 if hasattr(self, 'icon_pin_active'):
                     self.btn_pin.config(image=self.icon_pin_active)
             else:
                 if hasattr(self, 'icon_pin_inactive'):
                     self.btn_pin.config(image=self.icon_pin_inactive)
        elif isinstance(self.btn_pin, tk.Canvas):
             self.btn_pin.delete("all")
             color = self.colors.get("accent", "#007ACC") if self.config.pinned else self.colors.get("fg_dim", "#AAAAAA")
             self.btn_pin.create_oval(10, 5, 20, 15, fill=color, outline="")
             self.btn_pin.create_line(15, 15, 15, 25, fill=color, width=2)

        # Update Tooltip
        if self.pin_tooltip:
            self.pin_tooltip.text = "Unpin Window (Current: Pinned)" if self.config.pinned else "Pin Window (Current: Auto-Collapse)"

    def update_quick_create_icon(self, colors):
        """Updates Quick Create icon based on config actions."""
        if not self.btn_quick_create: return
        
        actions = self.config.quick_create_actions
        
        # Determine Color
        color = "#555555" # Grey (Disabled)
        if not actions:
            color = "#555555"
        elif len(actions) > 1:
            color = colors["fg_primary"]
        else:
            # Single Action Colors
            act = actions[0]
            if act == "New Email": color = "#6fb7ff"
            elif act == "New Meeting": color = "#ffb366"
            elif act == "New Appointment": color = "#ffcc80"
            elif act == "New Task": color = "#80e0a0"
            
        path = self.resource_path("icon2/plus.png")
        if os.path.exists(path):
            try:
                img = self.load_icon(path, size=(26, 26), color=color)
                self.icons["quick_create"] = img
                self.btn_quick_create.configure(image=img)
            except: pass

    def apply_theme(self, colors):
        """Updates all buttons with new theme colors."""
        self.colors = colors
        # 1. Header Buttons
        if self.btn_settings: self.btn_settings.config(bg=colors["bg_header"])
        if self.btn_help: self.btn_help.config(bg=colors["bg_header"], fg=colors["fg_dim"])
        if self.btn_refresh: self.btn_refresh.config(bg=colors["bg_header"])
        if self.btn_share: self.btn_share.config(bg=colors["bg_header"])
        
        # Re-load icons with new colors if needed
        # Settings
        self._reload_icon(self.btn_settings, "icon2/spanner.png", "settings", colors, (22,22))
        # Refresh
        self._reload_icon(self.btn_refresh, "icon2/refresh.png", "refresh", colors, (22,22))
        # Share
        self._reload_icon(self.btn_share, "icon2/share.png", "share", colors, (20,20))
        
        # Pin
        self._update_pin_theme(colors)
        
        # 2. Footer Buttons
        if self.lbl_version: self.lbl_version.config(bg=colors["bg_header"], fg=colors["fg_dim"])
        if self.btn_close: 
            self.btn_close.config(bg=colors["bg_header"])
            self._reload_icon(self.btn_close, "icon2/close-window.png", "close", colors, (28, 28), icon_color="#FF4444")
        if self.btn_quick_create: 
            self.btn_quick_create.config(bg=colors["bg_header"])
            self.update_quick_create_icon(colors)
            
        if self.btn_calendar: self.btn_calendar.config(bg=colors["bg_header"])
        self._reload_icon(self.btn_calendar, "icon2/calendar.png", "calendar", colors, (28,28))
        
        if self.btn_outlook: self.btn_outlook.config(bg=colors["bg_header"])
        self._reload_icon(self.btn_outlook, "icon2/email.png", "outlook", colors, (32,32))

    def _reload_icon(self, btn, path_rel, key, colors, size, icon_color=None):
        """Helper to re-load and apply an icon during theme change."""
        if not btn: return
        path = self.resource_path(path_rel)
        if os.path.exists(path):
            try:
                color = icon_color if icon_color else colors["fg_primary"]
                img = self.load_icon(path, size=size, color=color)
                self.icons[key] = img
                btn.config(image=img)
            except: pass

    def _update_pin_theme(self, colors):
        path = self.resource_path("icon2/pin1.png")
        if os.path.exists(path):
            try:
                self.icon_pin_active = self.load_icon(path, size=(24, 24), color=colors["accent"])
                self.icon_pin_inactive = self.load_icon(path, size=(24, 24), color=colors["fg_dim"])
                self.btn_pin.config(bg=colors["bg_header"])
                self.update_pin_state()
            except: pass
        elif isinstance(self.btn_pin, tk.Canvas):
            self.btn_pin.config(bg=colors["bg_header"])
            self._draw_pin_canvas(colors)
