import os
import sys
import win32com.client
from PIL import Image

def create_ico():
    try:
        if not os.path.exists("icons/app.ico"):
            print("Converting PNG to ICO...")
            img = Image.open("icons/Outlook_48x48.png")
            img.save("icons/app.ico", format='ICO', sizes=[(48, 48)])
            print("Created icons/app.ico")
    except Exception as e:
        print("Failed to create ICO: {}".format(e))

def create_shortcut():
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        path = os.path.join(desktop, "Outlook Monitor.lnk")
        
        # Target: pythonw.exe (no console) executing sidebar_main.py
        # We need absolute path to pythonw and the script
        python_exe = sys.executable.replace("python.exe", "pythonw.exe")
        script_path = os.path.abspath("sidebar_main.py")
        working_dir = os.path.dirname(script_path)
        icon_path = os.path.abspath("icons/app.ico")

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        shortcut.TargetPath = python_exe
        shortcut.Arguments = '"{}"'.format(script_path)
        shortcut.WorkingDirectory = working_dir
        shortcut.IconLocation = icon_path
        shortcut.Description = "Outlook Monitor Sidebar"
        shortcut.Save()
        
        print("Shortcut created at: {}".format(path))
        
    except Exception as e:
        print("Error creating shortcut: {}".format(e))

if __name__ == "__main__":
    create_ico()
    create_shortcut()
