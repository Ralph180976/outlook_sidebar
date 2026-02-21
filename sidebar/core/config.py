# -*- coding: utf-8 -*-
import sys
import os
from PIL import Image

# --- Resource Path Resolution (PyInstaller Support) ---
def resource_path(relative_path):
    """Resolve a relative path to an absolute path.
    
    When running as a PyInstaller bundle, files are extracted to sys._MEIPASS.
    When running normally, uses the script's directory as the base.
    """
    if getattr(sys, 'frozen', False):
        # Running as compiled exe (PyInstaller)
        base_path = sys._MEIPASS
    else:
        # Running as script — use the project root (parent of sidebar/)
        base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, relative_path)

# --- Application Constants ---
VERSION = "v1.3.16" # Flagged emails: filter fix, flag request + due date on hover

# --- Image Resampling Mode ---
try:
    # Pillow 10+
    RESAMPLE_MODE = Image.Resampling.LANCZOS
except AttributeError:
    # Older Pillow
    RESAMPLE_MODE = Image.ANTIALIAS

# --- Window configuration defaults ---
DEFAULT_MIN_WIDTH = 300
DEFAULT_HOT_STRIP_WIDTH = 16
DEFAULT_EXPANDED_WIDTH = 300
DEFAULT_FONT_FAMILY = "Segoe UI"
DEFAULT_FONT_SIZE = 9
