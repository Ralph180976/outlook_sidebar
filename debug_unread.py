import win32com.client
from datetime import datetime

def debug_unread():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("--- Checking Stores ---")
        for store in namespace.Stores:
            print("Store: {} (Type: {})".format(store.DisplayName, store.ExchangeStoreType))
            
            try:
                inbox = store.GetDefaultFolder(6) # olFolderInbox
                print("  Inbox: {} (Items: {}, Unread: {})".format(inbox.Name, inbox.Items.Count, inbox.UnReadItemCount))
                
                # Check actual items
                table = inbox.GetTable("[UnRead] = True")
                table.Columns.RemoveAll()
                table.Columns.Add("Subject")
                table.Columns.Add("ReceivedTime")
                table.Columns.Add("MessageClass")
                
                row_count = 0
                while not table.EndOfTable and row_count < 5:
                    row = table.GetNextRow()
                    if row:
                        vals = row.GetValues()
                        print("    - [Unread] Subject: '{}' (Class: {})".format(vals[0], vals[2]))
                        row_count += 1
                        
                if row_count == 0:
                     print("    - No unread items found via Table restrict.")
                     
            except Exception as e:
                print("  Error accessing inbox: {}".format(e))
                
    except Exception as e:
        print("Global Error: {}".format(e))

if __name__ == "__main__":
    debug_unread()
