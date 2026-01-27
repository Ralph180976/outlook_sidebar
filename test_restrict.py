import win32com.client

def test_restrict():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        print(f"Total items: {items.Count}")
        
        # Test IsMarkedAsTask
        try:
            flagged1 = items.Restrict("[IsMarkedAsTask] = True")
            print(f"Restrict('[IsMarkedAsTask] = True') count: {flagged1.Count}")
        except Exception as e:
            print(f"Restrict('[IsMarkedAsTask] = True') failed: {e}")
            
        # Test FlagStatus
        try:
            flagged2 = items.Restrict("[FlagStatus] = 2")
            print(f"Restrict('[FlagStatus] = 2') count: {flagged2.Count}")
        except Exception as e:
            print(f"Restrict('[FlagStatus] = 2') failed: {e}")
            
        # Test combined
        try:
            combined = items.Restrict("[UnRead] = True AND [FlagStatus] = 2")
            print(f"Restrict('[UnRead] = True AND [FlagStatus] = 2') count: {combined.Count}")
        except Exception as e:
            print(f"Restrict combined failed: {e}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_restrict()
