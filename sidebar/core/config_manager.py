
import json
import os
from sidebar.core.config import (
    DEFAULT_MIN_WIDTH, DEFAULT_EXPANDED_WIDTH,
    DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE
)

CONFIG_FILE = "sidebar_config.json"

class ConfigManager:
    """
    Centralized configuration manager.
    Enforces types and default values to prevent regression bugs.
    """
    def __init__(self):
        # Window Settings
        self.width = DEFAULT_EXPANDED_WIDTH
        self.pinned = True
        self.dock_side = "Left"
        self.theme = "Light"
        self.font_family = DEFAULT_FONT_FAMILY
        self.font_size = DEFAULT_FONT_SIZE
        self.window_mode = "dual"  # "single" or "dual"
        
        # Behavior
        self.poll_interval = 30
        self.hover_delay = 500
        self.hover_delay = 500
        self.show_hover_content = False
        self.email_double_click = False
        self.buttons_on_hover = False
        self.quick_create_actions = ["New Email"]
        
        # Email Filters
        self.show_read = False
        self.show_has_attachment = False
        self.only_flagged = False
        self.include_read_flagged = True
        self.flag_date_filter = "Anytime"
        
        # Email Content
        self.email_show_sender = True
        self.email_show_subject = True
        self.email_show_body = False
        self.email_body_lines = 2
        
        # Account Settings
        self.enabled_accounts = {} # {"Name": {"email": True, "calendar": True, ...}}
        
        # Reminder Filters (The source of previous bugs)
        self.reminder_show_flagged = True
        self.reminder_due_filters = ["No Date"] # List[str]
        self.reminder_show_categorized = True
        self.reminder_categories = []
        self.reminder_show_importance = False
        self.reminder_high_importance = False
        self.reminder_normal_importance = False
        self.reminder_low_importance = False
        
        self.reminder_show_meetings = True 
        self.reminder_pending_meetings = True
        self.reminder_accepted_meetings = True
        self.reminder_declined_meetings = False
        self.reminder_meeting_states = ["Accepted", "Tentative", "Appointments", "Received/Unknown"]
        self.reminder_meeting_dates = ["Today", "Tomorrow"] # List[str]
        self.reminder_custom_days = 30
        
        self.reminder_show_tasks = True
        self.reminder_tasks = True
        self.reminder_todo = True
        self.reminder_has_reminder = True
        self.reminder_task_dates = ["Overdue", "Today", "Tomorrow"] # List[str]
        
        # Buttons (Legacy support for now, moving to Toolbar module in Phase 2)
        self.btn_count = 2
        self.btn_config = [
            {"label": "Trash", "icon": "✓", "action1": "Mark Read", "action2": "Delete", "folder": ""}, 
            {"label": "Reply", "icon": "↩", "action1": "Reply", "action2": "None", "folder": ""}
        ]
        
        # Load immediately
        self.load()

    def load(self):
        """Loads config from disk, falling back to defaults for missing keys."""
        if not os.path.exists(CONFIG_FILE):
            return

        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                
            # Helper to safe-load with type enforcement could be added here
            # For now, distinct manual assignment prevents "typo leakage"
            
            self.width = data.get("width", self.width)
            self.pinned = data.get("pinned", self.pinned)
            self.dock_side = data.get("dock_side", self.dock_side)
            self.theme = data.get("theme", self.theme)
            self.font_family = data.get("font_family", self.font_family)
            self.font_size = data.get("font_size", self.font_size)
            self.show_hover_content = data.get("show_hover_content", self.show_hover_content)
            self.poll_interval = data.get("poll_interval", self.poll_interval)
            
            self.window_mode = data.get("window_mode", self.window_mode)
            
            self.show_read = data.get("show_read", self.show_read)
            self.show_has_attachment = data.get("show_has_attachment", self.show_has_attachment)
            self.only_flagged = data.get("only_flagged", self.only_flagged)
            self.include_read_flagged = data.get("include_read_flagged", self.include_read_flagged)
            self.flag_date_filter = data.get("flag_date_filter", self.flag_date_filter)
            
            self.enabled_accounts = data.get("enabled_accounts", self.enabled_accounts)
            
            # Reminders - STRICT naming
            # Handle legacy singluar 'reminder_due_filter' if present but prioritize list
            if "reminder_due_filters" in data:
                 self.reminder_due_filters = data.get("reminder_due_filters")
            elif "reminder_due_filter" in data:
                 # Auto-migration if old config exists
                 val = data.get("reminder_due_filter")
                 self.reminder_due_filters = [val] if isinstance(val, str) else val

            self.reminder_show_flagged = data.get("reminder_show_flagged", self.reminder_show_flagged)
            self.reminder_show_categorized = data.get("reminder_show_categorized", self.reminder_show_categorized)
            self.reminder_categories = data.get("reminder_categories", self.reminder_categories)
            
            self.reminder_show_importance = data.get("reminder_show_importance", self.reminder_show_importance)
            self.reminder_high_importance = data.get("reminder_high_importance", self.reminder_high_importance)
            self.reminder_normal_importance = data.get("reminder_normal_importance", self.reminder_normal_importance)
            self.reminder_low_importance = data.get("reminder_low_importance", self.reminder_low_importance)
            
            self.reminder_pending_meetings = data.get("reminder_pending_meetings", self.reminder_pending_meetings)
            self.reminder_accepted_meetings = data.get("reminder_accepted_meetings", self.reminder_accepted_meetings)
            self.reminder_declined_meetings = data.get("reminder_declined_meetings", self.reminder_declined_meetings)
            self.reminder_meeting_states = data.get("reminder_meeting_states", self.reminder_meeting_states)
            self.reminder_meeting_dates = data.get("reminder_meeting_dates", self.reminder_meeting_dates)
            
            self.reminder_tasks = data.get("reminder_tasks", self.reminder_tasks)
            self.reminder_todo = data.get("reminder_todo", self.reminder_todo)
            self.reminder_has_reminder = data.get("reminder_has_reminder", self.reminder_has_reminder)
            self.reminder_task_dates = data.get("reminder_task_dates", self.reminder_task_dates)
            
            self.buttons_on_hover = data.get("buttons_on_hover", self.buttons_on_hover)
            self.email_double_click = data.get("email_double_click", self.email_double_click)
            self.btn_count = data.get("btn_count", self.btn_count)
            self.btn_config = data.get("btn_config", self.btn_config)
            self.quick_create_actions = data.get("quick_create_actions", self.quick_create_actions)

        except Exception as e:
            print(f"Error loading config: {e}")

    def save(self):
        """Saves current state to disk."""
        data = {
            "width": self.width,
            "pinned": self.pinned,
            "dock_side": self.dock_side,
            "theme": self.theme,
            "font_family": self.font_family,
            "font_size": self.font_size,
            "show_hover_content": self.show_hover_content,
            "poll_interval": self.poll_interval,
            "window_mode": self.window_mode,
            
            "show_read": self.show_read,
            "show_has_attachment": self.show_has_attachment,
            "only_flagged": self.only_flagged,
            "include_read_flagged": self.include_read_flagged,
            "flag_date_filter": self.flag_date_filter,
            
            "enabled_accounts": self.enabled_accounts,
            
            "reminder_show_flagged": self.reminder_show_flagged,
            "reminder_due_filters": self.reminder_due_filters,
            "reminder_show_categorized": self.reminder_show_categorized,
            "reminder_categories": self.reminder_categories,
            "reminder_show_importance": self.reminder_show_importance,
            "reminder_high_importance": self.reminder_high_importance,
            "reminder_normal_importance": self.reminder_normal_importance,
            "reminder_low_importance": self.reminder_low_importance,
            
            "reminder_pending_meetings": self.reminder_pending_meetings,
            "reminder_accepted_meetings": self.reminder_accepted_meetings,
            "reminder_declined_meetings": self.reminder_declined_meetings,
            "reminder_meeting_states": self.reminder_meeting_states,
            "reminder_meeting_dates": self.reminder_meeting_dates,
            
            "reminder_tasks": self.reminder_tasks,
            "reminder_todo": self.reminder_todo,
            "reminder_has_reminder": self.reminder_has_reminder,
            "reminder_task_dates": self.reminder_task_dates,
            
            "buttons_on_hover": self.buttons_on_hover,
            "email_double_click": self.email_double_click,
            "btn_count": self.btn_count,
            "btn_config": self.btn_config,
            "quick_create_actions": self.quick_create_actions
        }
        
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")
