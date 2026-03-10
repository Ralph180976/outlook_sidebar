"""Test COM from different working directories"""
import os, sys

# Test 1: From dev directory
print("=== Test 1: CWD = Dev directory ===")
os.chdir(r"C:\Dev\Outlook_Sidebar")
print("CWD:", os.getcwd())
try:
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    stores = [s.DisplayName for s in ns.Stores]
    print("COM OK - {} stores: {}".format(len(stores), stores))
except Exception as e:
    print("COM FAILED: {}".format(e))

# Test 2: From installed directory
print("\n=== Test 2: CWD = Installed directory ===")
os.chdir(r"C:\Users\Ralph.SOUTHERNC\AppData\Local\InboxBar")
print("CWD:", os.getcwd())
try:
    outlook2 = win32com.client.Dispatch("Outlook.Application")
    ns2 = outlook2.GetNamespace("MAPI")
    stores2 = [s.DisplayName for s in ns2.Stores]
    print("COM OK - {} stores: {}".format(len(stores2), stores2))
except Exception as e:
    print("COM FAILED: {}".format(e))

# Test 3: Check if gen_py cache matters
print("\n=== Test 3: gen_py cache ===")
import win32com
print("win32com path:", win32com.__path__)
gen_py_path = os.path.join(os.environ.get("LOCALAPPDATA",""), "Temp", "gen_py")
print("gen_py exists:", os.path.exists(gen_py_path))
