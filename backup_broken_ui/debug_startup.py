
import sys
import os
import json

print("--- DIAGNOSTIC START ---")
print("Python Version: {}".format(sys.version))

print("\n[1] Checking Imports...")
try:
    import win32com.client
    print("SUCCESS: win32com.client")
except ImportError as e:
    print("FAIL: win32com.client - {}".format(e))

try:
    import tkinter
    print("SUCCESS: tkinter")
except ImportError as e:
    try:
        import Tkinter
        print("SUCCESS: Tkinter (Python 2)")
    except ImportError as e2:
        print("FAIL: tkinter/Tkinter - {}".format(e2))

try:
    from PIL import Image, ImageTk
    print("SUCCESS: PIL")
except ImportError as e:
    print("FAIL: PIL - {}".format(e))

print("\n[2] Checking Config...")
if os.path.exists("sidebar_config.json"):
    try:
        with open("sidebar_config.json", "r") as f:
            data = json.load(f)
            print("SUCCESS: Config loaded. Type: {}".format(type(data)))
    except Exception as e:
        print("FAIL: Config load error - {}".format(e))
else:
    print("WARNING: sidebar_config.json not found.")

print("\n[3] Checking Directories...")
if os.path.exists("icons"):
    print("SUCCESS: 'icons' directory exists.")
    print("Files in icons: {}".format(os.listdir("icons")[:10]))
else:
    print("FAIL: 'icons' directory missing.")

print("--- DIAGNOSTIC END ---")
