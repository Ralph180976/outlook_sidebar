import win32com.client

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

store = ns.Stores['ralph.marsh@coveya.co.uk']
inbox = store.GetDefaultFolder(6)

# Test 1: Can we read FlagRequest via Table API?
print("=== Table API - FlagRequest ===")
try:
    table = inbox.GetTable()
    table.Sort('ReceivedTime', True)
    table.Columns.RemoveAll()
    table.Columns.Add('Subject')
    table.Columns.Add('FlagStatus')
    try:
        table.Columns.Add('FlagRequest')
        print("FlagRequest column added OK")
    except Exception as e:
        print("FlagRequest NOT available in Table: {}".format(e))
    
    c = 0
    while not table.EndOfTable and c < 10:
        row = table.GetNextRow()
        vals = row.GetValues()
        flag_req = vals[2] if len(vals) > 2 else 'N/A'
        if vals[1] != 0 or flag_req:  # Show flagged or has FlagRequest
            print("  subj={}, status={}, request='{}'".format(vals[0][:50], vals[1], flag_req))
        c += 1
except Exception as e:
    print("Table error: {}".format(e))

# Test 2: Can we filter by FlagRequest?
print("\n=== Filter by FlagRequest <> '' ===")
try:
    table2 = inbox.GetTable("[FlagRequest] <> ''")
    table2.Columns.RemoveAll()
    table2.Columns.Add('Subject')
    table2.Columns.Add('FlagStatus')
    c = 0
    while not table2.EndOfTable and c < 10:
        row = table2.GetNextRow()
        vals = row.GetValues()
        print("  subj={}, status={}".format(vals[0][:50], vals[1]))
        c += 1
    print("  Total: {}".format(c))
except Exception as e:
    print("  Filter error: {}".format(e))

# Test 3: Check a non-flagged email's FlagRequest
print("\n=== Non-flagged email FlagRequest ===")
items = inbox.Items
items.Sort("[ReceivedTime]", True)
unflagged = items.Restrict("[FlagStatus] = 0")
for i in range(min(3, unflagged.Count)):
    item = unflagged.Item(i + 1)
    try:
        print("  subj={}, FlagRequest='{}'".format(item.Subject[:50], item.FlagRequest))
    except:
        print("  subj={}, FlagRequest=N/A".format(item.Subject[:50]))
