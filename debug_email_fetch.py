# Quick diagnostic: test if emails can be fetched
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from sidebar.services.outlook_client import OutlookClient

client = OutlookClient()
ok = client.connect()
print("Connected: {}".format(ok))
if not ok:
    print("FAILED to connect to Outlook")
    sys.exit(1)

print("\nAttempting to fetch 5 emails...")
try:
    emails, unread = client.get_inbox_items(count=5, unread_only=False)
    print("Got {} emails, {} unread".format(len(emails), unread))
    for i, e in enumerate(emails):
        print("  [{}] From: {} | Subject: {} | Preview: '{}'".format(
            i+1, e.get('sender','?'), e.get('subject','?')[:50], e.get('preview','')[:30]))
except Exception as ex:
    import traceback
    print("ERROR: {}".format(ex))
    traceback.print_exc()

# Also test a raw Table fetch to check GetValues works
print("\n--- Raw Table test ---")
try:
    for store in client.namespace.Stores:
        try:
            inbox = store.GetDefaultFolder(6)
            table = inbox.GetTable()
            table.Sort("ReceivedTime", True)
            table.Columns.RemoveAll()
            table.Columns.Add("EntryID")
            table.Columns.Add("Subject")
            table.Columns.Add("SenderName")
            table.Columns.Add("ReceivedTime")
            table.Columns.Add("UnRead")
            table.Columns.Add("FlagStatus")
            
            if not table.EndOfTable:
                row = table.GetNextRow()
                vals = row.GetValues()
                print("Store '{}': OK - {} columns, Subject='{}'".format(
                    store.DisplayName, len(vals), vals[1][:40] if vals[1] else "?"))
            else:
                print("Store '{}': Empty inbox".format(store.DisplayName))
        except Exception as se:
            print("Store '{}': ERROR - {}".format(store.DisplayName, se))
except Exception as ex:
    print("ERROR iterating stores: {}".format(ex))
