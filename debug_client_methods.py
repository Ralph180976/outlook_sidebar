
from sidebar_main import OutlookClient
import datetime

def debug_client():
    print("--- STARTING CLIENT DEBUG ---")
    try:
        client = OutlookClient()
        print("Client instantiated.")
        
        # Test get_accounts
        accounts = client.get_accounts()
        print("Accounts: {}".format(accounts))
        
        # Mock config
        config = {}
        for acc in accounts:
             config[acc] = {"email": True}

        # TEST: Manually call the helper for the first inbox found
        print("\n--- DEEP DIVE: _fetch_items_from_inbox_folder ---")
        try:
            store = client._get_enabled_stores([accounts[0]] if accounts else None).next()
            inbox = store.GetDefaultFolder(6)
            print("Target Folder: {}".format(inbox.Name))
            print("Unread Items: {}".format(inbox.UnReadItemCount))

            # DEBUGGING COLUMNS
            print("Debugging Table Columns...")
            table = inbox.GetTable()
            table.Columns.RemoveAll()
            
            cols_to_test = [
                ("EntryID", "EntryID"), 
                ("Subject", "Subject"),
                ("SenderName", "SenderName"),
                ("ReceivedTime", "ReceivedTime"),
                ("UnRead", "UnRead"),
                ("FlagStatus", "FlagStatus"),
                ("TaskDueDate", "TaskDueDate"),
                ("Importance", "Importance"),
                ("Categories", "Categories"),
                ("urn:schemas:httpmail:hasattachment", "HasAttach"),
                ("http://schemas.microsoft.com/mapi/proptag/0x1000001E", "Body")
            ]
            
            for col, name in cols_to_test:
                try:
                    table.Columns.Add(col)
                    print("  [SUCCESS] Added {}".format(name))
                    # Test fetch after EACH add to isolate the breaker
                    t_row = table.GetNextRow()
                    if t_row:
                        print("    -> Fetch verified")
                        table.MoveToStart() # Reset for next
                    else:
                         print("    -> WARNING: Table became empty after adding {}".format(name))
                except Exception as e:
                    print("  [FAILED] Adding {}: {}".format(name, e))
            
            # Fetch one row with what we have
            try:
                row = table.GetNextRow()
                if row:
                     print("  [SUCCESS] Row Fetch Success: {}".format(row.GetValues()))
                else:
                     print("  [WARNING] Table Empty after column mod")
            except Exception as e:
                 print("  [FAILED] Row Fetch Failed: {}".format(e))

            # Resuming Helper Call...
            print("Invoking original helper...")
            items = client._fetch_items_from_inbox_folder(
                folder=inbox,
                count=5,
                unread_only=False,
                only_flagged=False,
                due_filters=None,
                store=store
            )
            print("Helper Returned: {} items".format(len(items)))
            for i, x in enumerate(items):
                print("  [{}] {}".format(i, x.get("subject", "???")))
                
        except Exception as e:
            print("Deep dive failed: {}".format(e))
             
        print("\nCalling get_inbox_items...")
        items, count = client.get_inbox_items(
            count=10,
            unread_only=False, # Try false to ensure we get something
            account_names=accounts,
            account_config=config
        )
        
        print("Return Tuple: Items Type={}, Count={}".format(type(items), count))
        print("Item Count in List: {}".format(len(items)))
        
        if len(items) > 0:
            print("First Item Sample: {}".format(items[0]))
        else:
             print("WARNING: Item list is empty!")

    except Exception as e:
        print("CRITICAL CHECK FAILED: {}".format(e))
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_client()
