import win32com.client

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

store = ns.Stores['ralph.marsh@coveya.co.uk']
inbox = store.GetDefaultFolder(6)

# Use Items.Restrict instead of Table to get fresh data
items = inbox.Items
items.Sort("[ReceivedTime]", True)
flagged = items.Restrict("[FlagStatus] <> 0")

print("=== Items.Restrict approach ===")
for i in range(min(flagged.Count, 10)):
    item = flagged.Item(i + 1)
    print("  subj={}, FlagStatus={}, FlagIcon={}".format(
        item.Subject, item.FlagStatus, item.FlagIcon if hasattr(item, 'FlagIcon') else 'N/A'))

print()
print("=== Table approach ===")
table = inbox.GetTable('[FlagStatus] <> 0')
table.Columns.RemoveAll()
table.Columns.Add('Subject')
table.Columns.Add('FlagStatus')
c = 0
while not table.EndOfTable and c < 10:
    row = table.GetNextRow()
    vals = row.GetValues()
    print("  subj={}, FlagStatus={}".format(vals[0], vals[1]))
    c += 1
if c == 0:
    print("  (none)")
