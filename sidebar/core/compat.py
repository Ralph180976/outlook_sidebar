# -*- coding: utf-8 -*-
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
