
import win32com.client

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    print("--- Polling Stores ---")
    for store in namespace.Stores:
        print(f"Store: {store.DisplayName} (Type: {store.ExchangeStoreType})")
        
        try:
            # Try to get Inbox (6)
            inbox = store.GetDefaultFolder(6)
            print(f"  - Inbox: {inbox.Name} (Path: {inbox.FolderPath})")
        except Exception as e:
            print(f"  - Inbox N/A: {e}")
            
        try:
             # Try to get Calendar (9)
            calendar = store.GetDefaultFolder(9)
            print(f"  - Calendar: {calendar.Name} (Path: {calendar.FolderPath})")
        except Exception as e:
            print(f"  - Calendar N/A: {e}")
            
except Exception as e:
    print(f"Error: {e}")
