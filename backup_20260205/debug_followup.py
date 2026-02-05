import win32com.client
from datetime import datetime

def debug_followup_dates():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        # We only care about flagged items for this research
        flagged_items = items.Restrict("[FlagStatus] <> 0")
        
        print(f"Flagged items found: {flagged_items.Count}")
        print("-" * 60)
        
        for item in flagged_items:
            print(f"Subject: {item.Subject}")
            # Standard properties
            print(f"  TaskStartDate: {getattr(item, 'TaskStartDate', 'N/A')}")
            print(f"  TaskDueDate: {getattr(item, 'TaskDueDate', 'N/A')}")
            print(f"  TaskCompletedDate: {getattr(item, 'TaskCompletedDate', 'N/A')}")
            print(f"  FlagRequest: {getattr(item, 'FlagRequest', 'N/A')}")
            print(f"  ToDoItemParent: {getattr(item, 'ToDoItemParent', 'N/A')}")
            
            # Use PropertyAccessor for MAPI properties if needed
            pa = item.PropertyAccessor
            # PR_FOLLOWUP_ICON (0x10950003)
            # PR_TODO_ITEM_FLAGS (0x0E2B0003)
            # PR_FLAG_COMPLETE_TIME (0x10910040)
            
            mapi_props = {
                "PR_FOLLOWUP_ICON": "http://schemas.microsoft.com/mapi/proptag/0x10950003",
                "PR_TASK_DUE_DATE": "http://schemas.microsoft.com/mapi/proptag/0x81050040", # Often used for tasks
                "PR_REPLY_TIME": "http://schemas.microsoft.com/mapi/proptag/0x00300040"
            }
            
            for name, uri in mapi_props.items():
                try:
                    val = pa.GetProperty(uri)
                    print(f"  {name}: {val}")
                except:
                    pass
            print("-" * 40)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug_followup_dates()
