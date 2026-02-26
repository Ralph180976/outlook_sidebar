# -*- coding: utf-8 -*-
"""
InboxBar Diagnostic Tool
========================
Run this script to diagnose Outlook connection issues.
It will check each component and report what's working and what isn't.

Usage: Double-click this file, or run: py -3 diagnose.py
"""
import sys
import os
import traceback

print("=" * 60)
print("  InboxBar Diagnostic Tool")
print("=" * 60)
print()

# 1. Python Version
print("[1] Python Version")
print(f"    Version: {sys.version}")
print(f"    Executable: {sys.executable}")
print(f"    Platform: {sys.platform}")
print()

# 2. Required Modules
print("[2] Required Modules")
modules = {
    "win32com.client": "pywin32 (Outlook COM)",
    "pythoncom": "pywin32 (COM threading)",
    "tkinter": "UI framework",
    "PIL": "Pillow (image processing)",
}
for mod, desc in modules.items():
    try:
        __import__(mod)
        print(f"    ✓ {mod} ({desc})")
    except ImportError:
        print(f"    ✗ {mod} ({desc}) — NOT INSTALLED")
        print(f"      Fix: pip install {desc.split('(')[0].strip()}")
print()

# 3. Outlook Profile Check
print("[3] Outlook Profile Check")
try:
    import winreg
    key_path = r"Software\Microsoft\Office\16.0\Outlook\Profiles"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
        count = 0
        while True:
            try:
                name = winreg.EnumKey(key, count)
                print(f"    ✓ Profile found: {name}")
                count += 1
            except OSError:
                break
        winreg.CloseKey(key)
        if count == 0:
            print("    ⚠ No Outlook profiles found in registry")
    except FileNotFoundError:
        # Try Office 15.0 (Outlook 2013)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\15.0\Outlook\Profiles")
            print("    ✓ Found Outlook 2013 profiles")
            winreg.CloseKey(key)
        except:
            print("    ⚠ No Outlook profile registry keys found")
            print("      This may mean Outlook has never been set up on this machine")
except Exception as e:
    print(f"    ✗ Registry check failed: {e}")
print()

# 4. Outlook COM Connection
print("[4] Outlook COM Connection")
outlook_app = None
namespace = None
try:
    import win32com.client
    import pythoncom
    
    # Try Dispatch first (connects to running Outlook)
    print("    Attempting win32com.client.Dispatch('Outlook.Application')...")
    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        print(f"    ✓ Outlook Application object created")
        print(f"    ✓ Outlook Version: {outlook_app.Version}")
    except Exception as e:
        print(f"    ✗ Dispatch failed: {e}")
        print("    Attempting win32com.client.gencache.EnsureDispatch...")
        try:
            outlook_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            print(f"    ✓ EnsureDispatch succeeded")
        except Exception as e2:
            print(f"    ✗ EnsureDispatch also failed: {e2}")
    
    if outlook_app:
        print()
        print("    Attempting GetNamespace('MAPI')...")
        try:
            namespace = outlook_app.GetNamespace("MAPI")
            print(f"    ✓ MAPI Namespace connected")
        except Exception as e:
            print(f"    ✗ GetNamespace failed: {e}")
            traceback.print_exc()
except Exception as e:
    print(f"    ✗ COM initialization failed: {e}")
    traceback.print_exc()
print()

# 5. Stores / Accounts
print("[5] Mail Stores (Accounts)")
if namespace:
    try:
        store_count = 0
        for store in namespace.Stores:
            try:
                name = store.DisplayName
                store_id = store.StoreID[:30] + "..."
                print(f"    ✓ Store: {name}")
                
                # Try getting inbox
                try:
                    inbox = store.GetDefaultFolder(6)
                    unread = inbox.UnReadItemCount
                    total = inbox.Items.Count
                    print(f"      Inbox: {total} items, {unread} unread")
                except Exception as e:
                    print(f"      ⚠ Could not access Inbox: {e}")
                
                store_count += 1
            except Exception as e:
                print(f"    ⚠ Error reading store: {e}")
        
        if store_count == 0:
            print("    ⚠ No stores found!")
    except Exception as e:
        print(f"    ✗ Error enumerating stores: {e}")
        traceback.print_exc()
else:
    print("    ✗ Skipped (no MAPI connection)")
print()

# 6. Test Table API (most common failure point)
print("[6] Table API Test")
if namespace:
    try:
        for store in namespace.Stores:
            try:
                inbox = store.GetDefaultFolder(6)
                table = inbox.GetTable()
                table.Columns.RemoveAll()
                table.Columns.Add("Subject")
                if not table.EndOfTable:
                    row = table.GetNextRow()
                    vals = row.GetValues()
                    print(f"    ✓ Table API works — first email: '{vals[0][:50]}...'")
                else:
                    print(f"    ✓ Table API works (inbox is empty)")
                break
            except Exception as e:
                continue
    except Exception as e:
        print(f"    ✗ Table API failed: {e}")
else:
    print("    ✗ Skipped (no MAPI connection)")
print()

# 7. File System Check
print("[7] InboxBar Files Check")
script_dir = os.path.dirname(os.path.abspath(__file__))
critical_files = [
    "sidebar_main.py",
    "sidebar/__init__.py",
    "sidebar/services/outlook_client.py",
    "sidebar/core/config_manager.py",
]
for f in critical_files:
    full = os.path.join(script_dir, f)
    if os.path.exists(full):
        print(f"    ✓ {f}")
    else:
        print(f"    ✗ {f} — MISSING")

# Check config
config_path = os.path.join(os.environ.get("LOCALAPPDATA", "."), "OutlookSidebar", "config.json")
if os.path.exists(config_path):
    print(f"    ✓ Config: {config_path}")
    try:
        import json
        with open(config_path, "r") as fp:
            cfg = json.load(fp)
        print(f"      Backend: {cfg.get('backend', 'not set')}")
        print(f"      Theme: {cfg.get('theme', 'not set')}")
        ea = cfg.get("enabled_accounts", {})
        if ea:
            print(f"      Accounts configured: {list(ea.keys())}")
        else:
            print(f"      Accounts: Not configured yet")
    except Exception as e:
        print(f"      ⚠ Could not read config: {e}")
else:
    print(f"    ℹ Config not found (first run): {config_path}")
print()

# 8. Summary
print("=" * 60)
print("  SUMMARY")
print("=" * 60)
issues = []
if not outlook_app:
    issues.append("Outlook COM connection FAILED — is Outlook running?")
if not namespace:
    issues.append("MAPI namespace FAILED — Outlook may not be fully started")

if issues:
    print()
    print("  ⚠ ISSUES FOUND:")
    for i, issue in enumerate(issues, 1):
        print(f"    {i}. {issue}")
    print()
    print("  SUGGESTED FIXES:")
    print("    1. Make sure Outlook is fully open and loaded")
    print("    2. Close and reopen Outlook, then try again")
    print("    3. Run InboxBar as the same user that owns Outlook")
    print("    4. Check if antivirus is blocking COM access")
else:
    print()
    print("  ✓ All checks passed! Outlook connection is working.")
    print()
    print("  If InboxBar still isn't working, please share the")
    print("  log file at:")
    log_path = os.path.join(os.environ.get("LOCALAPPDATA", "."), "OutlookSidebar", "debug_outlook.log")
    print(f"    {log_path}")

print()
input("Press Enter to close...")
