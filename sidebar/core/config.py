# -*- coding: utf-8 -*-
import sys
import os
from PIL import Image

# --- Application Constants ---
VERSION = "v1.3.6" # Reminder button layout fix & icon updates

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
