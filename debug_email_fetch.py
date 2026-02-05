
import win32com.client
import datetime
from datetime import timedelta     

def debug_fetch():
    print("--- STARTING DEBUG FETCH ---")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("Connected to Outlook NameSpace")
    except Exception as e:
        print("CRITICAL: Failed to connect to Outlook: {}".format(e))
        return

    try:
        stores = namespace.Stores
        print("Found {} stores.".format(stores.Count))
        
        for store in stores:
            print("\nScanning Store: {}".format(store.DisplayName))
            try:
                inbox = store.GetDefaultFolder(6) # olFolderInbox
                print("  Inbox: {}".format(inbox.Name))
                print("  Unread Count: {}".format(inbox.UnReadItemCount))
                
                # Try Table access (sidebar method)
                table = inbox.GetTable()
                count = table.GetRowCount()
                print("  Table Row Count: {}".format(count))
                
                # Try fetching 5 rows
                table.Columns.RemoveAll()
                table.Columns.Add("Subject")
                table.Columns.Add("ReceivedTime")
                
                rows_fetched = 0
                while not table.EndOfTable and rows_fetched < 5:
                    row = table.GetNextRow()
                    vals = row.GetValues()
                    print("    Msg {}: {} ({})".format(rows_fetched+1, vals[0], vals[1]))
                    rows_fetched += 1
                    
            except Exception as e:
                print("  Error scanning store: {}".format(e))
                
    except Exception as e:
        print("Error iterating stores: {}".format(e))

    print("--- END DEBUG FETCH ---")

if __name__ == "__main__":
    debug_fetch()
