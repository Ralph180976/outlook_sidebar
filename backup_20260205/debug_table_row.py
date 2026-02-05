import win32com.client
import datetime

def test_row_access():
    try:
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
        
        # Simple GetTable
        table = inbox.GetTable()
        table.Columns.RemoveAll()
        
        # Add a few columns, including complex properites
        cols = ["EntryID", "Subject", "SenderName"]
        for c in cols:
            table.Columns.Add(c)
            
        attach_prop = "urn:schemas:httpmail:hasattachment"
        try: table.Columns.Add(attach_prop)
        except: print("Fail Add Attach")

        body_prop = "http://schemas.microsoft.com/mapi/proptag/0x1000001E"
        try: table.Columns.Add(body_prop)
        except: print("Fail Add Body")
        
        print(f"Total Columns: {table.Columns.Count}")
        for i in range(1, table.Columns.Count + 1):
             print(f"Col {i}: {table.Columns(i).Name}")
        
        # Sort and get first row
        table.Sort("ReceivedTime", True)
        row = table.GetNextRow()
        
        if row:
            print("\nRow retrieved.")
            
            # Test Access - String Key if pywin32 supports dict access
            try:
                print(f"Access ['Subject']: {row['Subject']}")
            except Exception as e:
                print(f"FAIL ['Subject']: {e}")
                
            # Test Access - Attribute
            try:
                print(f"Access .Subject: {row.Subject}")
            except Exception as e:
                print(f"FAIL .Subject: {e}")
                
            # Test Access - Item Method (1-based index)
            try:
                print(f"Access Item(2): {row.Item(2)}") # Subject should be 2?
            except Exception as e:
                print(f"FAIL Item(2): {e}")
                
            # Test Access - Complex Prop Name String Key
            try:
                print(f"Access ['{body_prop}']: {row[body_prop]}")
            except Exception as e:
                print(f"FAIL ['BodyPropString']: {e}")
                
            # Test Access - Using the actual Column Name that was added?
            # Maybe the column name is simplified?
            # Or maybe we need to iterate columns to find the index for the property?
            
            # Let's check `GetValues()` if Row has it
            try:
                 print(f"GetValues(): {row.GetValues()}")
            except: 
                 pass
                 
    except Exception as e:
        print(f"Critical Fail: {e}")

if __name__ == "__main__":
    test_row_access()
