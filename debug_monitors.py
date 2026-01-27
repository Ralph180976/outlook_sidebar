import tkinter as tk
import ctypes
from ctypes import wintypes

user32 = ctypes.windll.user32

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('rcMonitor', wintypes.RECT),
        ('rcWork', wintypes.RECT),
        ('dwFlags', wintypes.DWORD),
    ]

def get_monitors():
    monitors = []
    def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi)):
            monitors.append({
                "hMonitor": hMonitor,
                "rcMonitor": (mi.rcMonitor.left, mi.rcMonitor.top, mi.rcMonitor.right, mi.rcMonitor.bottom),
                "rcWork": (mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom),
                "flags": mi.dwFlags
            })
        return True

    MONITORENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.ULONG, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)
    user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0)
    return monitors

def debug_window_info():
    root = tk.Tk()
    root.title("Monitor Debugger")
    root.geometry("400x300")
    
    text = tk.Text(root)
    text.pack(fill="both", expand=True)
    
    def refresh():
        text.delete(1.0, tk.END)
        text.insert(tk.END, f"Screen Width: {root.winfo_screenwidth()}\n")
        text.insert(tk.END, f"Screen Height: {root.winfo_screenheight()}\n\n")
        
        ms = get_monitors()
        for i, m in enumerate(ms):
            text.insert(tk.END, f"Monitor {i}:\n")
            text.insert(tk.END, f"  Handle: {m['hMonitor']}\n")
            text.insert(tk.END, f"  Rect: {m['rcMonitor']}\n")
            text.insert(tk.END, f"  Work: {m['rcWork']}\n")
            text.insert(tk.END, f"  Flags: {m['flags']}\n\n")
            
        # Get current monitor for this window
        hwnd = root.winfo_id()
        hMon = user32.MonitorFromWindow(hwnd, 2) # MONITOR_DEFAULTTONEAREST
        text.insert(tk.END, f"Current Window Monitor Handle: {hMon}\n")
        
    btn = tk.Button(root, text="Refresh Info", command=refresh)
    btn.pack(side="bottom", fill="x")
    
    refresh()
    root.mainloop()

if __name__ == "__main__":
    debug_window_info()
