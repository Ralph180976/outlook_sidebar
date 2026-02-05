import win32com.client

def test_schema_restrict():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        print(f"Total items: {items.Count}")
        
        # PR_FLAG_STATUS = 0x10900003
        # olFlagMarked = 2
        schema_query = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
        
        try:
            # Note: Restrict syntax for schema properties can be specific
            flagged = items.Restrict(f'"{schema_query}" = 2')
            print(f"Restrict schema query count: {flagged.Count}")
            if flagged.Count > 0:
                print(f"First flagged item: {flagged.GetFirst().Subject}")
        except Exception as e:
            print(f"Restrict schema query failed: {e}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_schema_restrict()
