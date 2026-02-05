import win32com.client

def test_flag_restrict_variations():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        print(f"Total items: {items.Count}")
        
        variations = [
            "[FlagStatus] = 2",
            "[FlagStatus] <> 0",
            "[FlagIcon] = 6",
            "[TaskStatus] = 0",
            "[Importance] = 2", # Just to check if Restrict works at all
            "NOT([FlagStatus] = 0)"
        ]
        
        for v in variations:
            try:
                res = items.Restrict(v)
                print(f"Restrict('{v}') count: {res.Count}")
            except Exception as e:
                print(f"Restrict('{v}') failed: {e}")
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_flag_restrict_variations()
