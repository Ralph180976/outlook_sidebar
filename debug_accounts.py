import win32com.client

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    print("--- Polling Stores ---")
    for store in namespace.Stores:
        print("Store: {} (Type: {})".format(store.DisplayName, store.ExchangeStoreType))
        
        try:
            # Try to get Inbox (6)
            inbox = store.GetDefaultFolder(6)
            print("  - Inbox: {} (Path: {})".format(inbox.Name, inbox.FolderPath))
        except Exception as e:
            print("  - Inbox N/A: {}".format(e))
            
        try:
             # Try to get Calendar (9)
            calendar = store.GetDefaultFolder(9)
            print("  - Calendar: {} (Path: {})".format(calendar.Name, calendar.FolderPath))
        except Exception as e:
            print("  - Calendar N/A: {}".format(e))
            
except Exception as e:
    print("Error: {}".format(e))
