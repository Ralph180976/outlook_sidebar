import win32com.client

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

store = ns.Stores['ralph.marsh@coveya.co.uk']
inbox = store.GetDefaultFolder(6)

# Table API check
print("=== Table API ===")
table = inbox.GetTable('[FlagStatus] <> 0')
table.Columns.RemoveAll()
table.Columns.Add('Subject')
table.Columns.Add('FlagRequest')
try:
    table.Columns.Add('TaskDueDate')
    print("TaskDueDate column added OK")
except Exception as e:
    print("TaskDueDate NOT available: {}".format(e))

c = 0
while not table.EndOfTable and c < 10:
    row = table.GetNextRow()
    vals = row.GetValues()
    print("  subj={}, req='{}', due='{}'  type={}".format(
        vals[0][:40], vals[1], vals[2] if len(vals) > 2 else 'N/A',
        type(vals[2]).__name__ if len(vals) > 2 else 'N/A'))
    c += 1

# Items API check for comparison
print("\n=== Items API ===")
items = inbox.Items
flagged = items.Restrict("[FlagStatus] <> 0")
for i in range(min(flagged.Count, 10)):
    item = flagged.Item(i + 1)
    try:
        due = item.TaskDueDate
        print("  subj={}, TaskDueDate={}, year={}".format(
            item.Subject[:40], due, due.year if due else 'None'))
    except Exception as e:
        print("  subj={}, TaskDueDate error: {}".format(item.Subject[:40], e))
