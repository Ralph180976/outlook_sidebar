# -*- coding: utf-8 -*-

# --- Outlook Category Colors (Approximate Hex for Dark Mode) ---
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

# --- Application Theme Palettes ---
COLOR_PALETTES = {
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
        "accent": "#007ACC",
        "divider": "#D0D0D0",
        "scroll_bg": "#E8E8E8",
        "input_bg": "#FFFFFF",
        "card_border": "#E5E5E5"
    }
}
