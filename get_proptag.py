import win32com.client

def get_proptag():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        for item in items:
            if getattr(item, "FlagStatus", 0) == 2:
                print(f"Subject: {item.Subject}")
                pa = item.PropertyAccessor
                
                # Try to find common flagging/task properties
                props = [
                    "http://schemas.microsoft.com/mapi/proptag/0x10900003", # PR_FLAG_STATUS
                    "http://schemas.microsoft.com/mapi/proptag/0x0E060003", # PR_MESSAGE_FLAGS
                    "http://schemas.microsoft.com/mapi/proptag/0x3001001F", # PR_DISPLAY_NAME
                ]
                
                for p in props:
                    try:
                        val = pa.GetProperty(p)
                        print(f"  {p}: {val}")
                    except Exception as e:
                        print(f"  {p} failed: {e}")
                
                # Also try Task-specific properties if any
                break
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_proptag()
