# -*- coding: utf-8 -*-
import ctypes
from ctypes import wintypes
import sys

# --- Windows API Constants & Structures ---
ABM_NEW = 0x00000000
ABM_REMOVE = 0x00000001
ABM_QUERYPOS = 0x00000002
ABM_SETPOS = 0x00000003
ABM_GETSTATE = 0x00000004
ABM_GETTASKBARPOS = 0x00000005
ABM_ACTIVATE = 0x00000006
ABM_GETAUTOHIDEBAR = 0x00000007
ABM_SETAUTOHIDEBAR = 0x00000008
ABM_WINDOWPOSCHANGED = 0x00000009
ABM_SETSTATE = 0x0000000A

ABE_LEFT = 0
ABE_TOP = 1
ABE_RIGHT = 2
ABE_BOTTOM = 3

# Window Messages
WM_ACTIVATE = 0x0006
WM_WINDOWPOSCHANGED = 0x0047
GWLP_WNDPROC = -4

SWP_NOACTIVATE = 0x0010
SWP_NOZORDER = 0x0004
SWP_SHOWWINDOW = 0x0040
SWP_NOSENDCHANGING = 0x0400

# Callback Event ID
ABN_POSCHANGED = 0x00000001

class APPBARDATA(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('hWnd', wintypes.HWND),
        ('uCallbackMessage', wintypes.UINT),
        ('uEdge', wintypes.UINT),
        ('rc', wintypes.RECT),
        ('lParam', wintypes.LPARAM),
    ]

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('rcMonitor', wintypes.RECT),
        ('rcWork', wintypes.RECT),
        ('dwFlags', wintypes.DWORD),
    ]

# Setup PLL
shell32 = ctypes.windll.shell32
user32 = ctypes.windll.user32

# Define precise argtypes for 64-bit safety
shell32.SHAppBarMessage.argtypes = (wintypes.DWORD, ctypes.POINTER(APPBARDATA))
shell32.SHAppBarMessage.restype = wintypes.UINT # UINT_PTR in docs, but UINT sufficient for return logic usually

user32.SetWindowPos.argtypes = (
    wintypes.HWND, wintypes.HWND,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.UINT
)
user32.SetWindowPos.restype = wintypes.BOOL

# WNDPROC Type
WNDPROC = ctypes.WINFUNCTYPE(
    wintypes.LPARAM, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
)
# Note: LRESULT == LPARAM (LONG_PTR) in ctypes for practical purposes on 64-bit

class AppBarManager:
    """
    Manages the Windows AppBar registration and positioning.
    Includes WndProc hooking for proper system integration.
    """
    def __init__(self, hwnd, edge=ABE_LEFT):
        self.hwnd = hwnd
        self.edge = edge
        self.registered = False
        self.uCallbackMessage = 0x0400 + 1  # WM_USER + 1
        
        # Keep appbar data persistent
        self.abd = APPBARDATA()
        self.abd.cbSize = ctypes.sizeof(APPBARDATA)
        self.abd.hWnd = self.hwnd
        self.abd.uCallbackMessage = self.uCallbackMessage
        self.abd.uEdge = self.edge
        
        # State storage for re-entrant calls
        self._last_width = 0
        self._mon_left = 0
        self._mon_top = 0
        self._mon_w = 0
        self._mon_h = 0
        
        # Hook references
        self._proc = None
        self._old_proc = None
        self._call_old = None

    def register(self):
        """Registers the window as an AppBar."""
        if not self.registered:
            shell32.SHAppBarMessage(ABM_NEW, ctypes.byref(self.abd))
            self.registered = True

    def unregister(self):
        """Unregisters the AppBar."""
        if self.registered:
            shell32.SHAppBarMessage(ABM_REMOVE, ctypes.byref(self.abd))
            self.registered = False
            
    def set_pos(self, width, monitor_left, monitor_top, monitor_width, monitor_height):
        """
        Reserving space for the AppBar on the current monitor's edge.
        """
        if not self.registered:
            return

        # Store for callback usage
        self._last_width = width
        self._mon_left = monitor_left
        self._mon_top = monitor_top
        self._mon_w = monitor_width
        self._mon_h = monitor_height

        # 1. Query Position
        if self.edge == ABE_LEFT:
            self.abd.rc.left = monitor_left
            self.abd.rc.top = monitor_top
            self.abd.rc.right = monitor_left + width
            self.abd.rc.bottom = monitor_top + monitor_height
        elif self.edge == ABE_RIGHT:
            self.abd.rc.left = monitor_left + monitor_width - width
            self.abd.rc.top = monitor_top
            self.abd.rc.right = monitor_left + monitor_width
            self.abd.rc.bottom = monitor_top + monitor_height
        
        # Query the system for an approved position
        shell32.SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(self.abd))
        
        # 2. Adjust if necessary (System might have changed it)
        if self.edge == ABE_LEFT:
            self.abd.rc.right = self.abd.rc.left + width
        elif self.edge == ABE_RIGHT:
            self.abd.rc.left = self.abd.rc.right - width
        
        # 3. Set Position
        shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(self.abd))
        
        # 4. Force Window Position to match AppBar reservation
        # Fix for "other windows going behind" - we must be exactly where the system tells us
        rc = self.abd.rc
        w = rc.right - rc.left
        h = rc.bottom - rc.top
        user32.SetWindowPos(
            self.hwnd, None, 
            rc.left, rc.top, w, h,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING
        )
        
        # Return the actual rectangle committed
        return rc.left, rc.top, w, h
        
    def hook_wndproc(self):
        """Installs the WndProc hook to handle AppBar messages."""
        if self._proc:
            return # Already hooked

        # Define the hook implementation
        def wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_ACTIVATE:
                shell32.SHAppBarMessage(ABM_ACTIVATE, ctypes.byref(self.abd))

            elif msg == WM_WINDOWPOSCHANGED:
                shell32.SHAppBarMessage(ABM_WINDOWPOSCHANGED, ctypes.byref(self.abd))

            elif msg == self.uCallbackMessage:
                # Explorer notifies us that something changed (taskbar/appbar/monitor/etc.)
                if wParam == ABN_POSCHANGED:
                    # Re-apply position
                    if self._last_width > 0:
                        self.set_pos(
                            self._last_width,
                            self._mon_left, self._mon_top,
                            self._mon_w, self._mon_h
                        )
                return 0

            return self._call_old(hWnd, msg, wParam, lParam)

        # Create C-callable function pointer
        # casting logic is handled by decorated function in standard python ctypes if done right, 
        # but here we use WNDPROC(...) manually.
        self._proc = WNDPROC(wnd_proc)
        
        # Install Hook
        self._old_proc, self._call_old = _set_window_proc(self.hwnd, self._proc)

# Helper for 64-bit safe SetWindowLong
def _set_window_proc(hwnd, new_proc):
    is_64 = ctypes.sizeof(ctypes.c_void_p) == 8
    
    # We need to handle result type carefully
    
    if is_64:
        SetWindowLongPtr = user32.SetWindowLongPtrW
        SetWindowLongPtr.argtypes = (wintypes.HWND, ctypes.c_int, ctypes.c_void_p)
        SetWindowLongPtr.restype = ctypes.c_void_p
        
        GetWindowLongPtr = user32.GetWindowLongPtrW
        GetWindowLongPtr.argtypes = (wintypes.HWND, ctypes.c_int)
        GetWindowLongPtr.restype = ctypes.c_void_p
        
        old_ptr = GetWindowLongPtr(hwnd, GWLP_WNDPROC)
        new_ptr = ctypes.cast(new_proc, ctypes.c_void_p)
        SetWindowLongPtr(hwnd, GWLP_WNDPROC, new_ptr)
        old_proc = old_ptr
    else:
        SetWindowLong = user32.SetWindowLongW
        SetWindowLong.argtypes = (wintypes.HWND, ctypes.c_int, ctypes.c_long)
        SetWindowLong.restype = ctypes.c_long
        
        GetWindowLong = user32.GetWindowLongW
        GetWindowLong.argtypes = (wintypes.HWND, ctypes.c_int)
        GetWindowLong.restype = ctypes.c_long
        
        # On 32-bit, callbacks are passed as memory addresses (int/long)
        new_ptr = ctypes.cast(new_proc, ctypes.c_void_p).value
        # Note: ctypes WINFUNCTYPE instance can be cast to void_p
        
        old_proc = GetWindowLong(hwnd, GWLP_WNDPROC)
        SetWindowLong(hwnd, GWLP_WNDPROC, new_ptr)

    # Prepare CallWindowProc
    user32.CallWindowProcW.argtypes = (ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    user32.CallWindowProcW.restype = wintypes.LPARAM

    def call_old(hWnd, msg, wParam, lParam):
        return user32.CallWindowProcW(old_proc, hWnd, msg, wParam, lParam)

    return old_proc, call_old
