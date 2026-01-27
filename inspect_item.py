import win32com.client

def inspect_flagged_item():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        for item in items:
            if getattr(item, "IsMarkedAsTask", False) == True or getattr(item, "FlagStatus", 0) == 2:
                print(f"Subject: {item.Subject}")
                print(f"  FlagStatus: {item.FlagStatus}")
                print(f"  IsMarkedAsTask: {item.IsMarkedAsTask}")
                print(f"  ToDoItemParent: {getattr(item, 'ToDoItemParent', 'N/A')}")
                
                # Check for UserProperties or other specific tags if possible
                # But let's try to see if there's a simple one.
                
                # Try to use the UserProperty names if any
                break
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_flagged_item()
