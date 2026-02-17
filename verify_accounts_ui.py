# -*- coding: utf-8 -*-
import sys
try:
    import Tkinter as tk
except ImportError:
    import tkinter as tk

from sidebar.ui.panels.account_settings import AccountSelectionUI

def mock_folder_selector(account, callback):
    print("Mock: Opening folder selector for {}".format(account))
    callback("Inbox/Test")

def verify_accounts():
    root = tk.Tk()
    root.title("Account Selection UI Verification")
    root.geometry("400x600")
    
    accounts = ["user@example.com", "other@example.com"]
    enabled = {"user@example.com": {"email": True, "calendar": True}}
    
    print("--- Testing AccountSelectionUI ---")
    try:
        ui = AccountSelectionUI(root, accounts, enabled, mock_folder_selector)
        ui.pack(fill="both", expand=True)
        print("SUCCESS: AccountSelectionUI instantiated")
        
        # Keep window open briefly or just check instantiation?
        # Let's run mainloop briefly to let it render? No, just instantiation check.
        root.update()
        print("SUCCESS: AccountSelectionUI rendered")
        
    except Exception as e:
        print("FAILURE: AccountSelectionUI failed: {}".format(e))
        import traceback
        traceback.print_exc()

    root.destroy()

if __name__ == "__main__":
    verify_accounts()
