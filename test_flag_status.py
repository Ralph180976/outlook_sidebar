import win32com.client

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

for store in ns.Stores:
    try:
        inbox = store.GetDefaultFolder(6)
        # Check ALL items with any flag (status 1 or 2)
        for flag_val in [1, 2]:
            table = inbox.GetTable('[FlagStatus] = {}'.format(flag_val))
            table.Columns.RemoveAll()
            table.Columns.Add('Subject')
            table.Columns.Add('FlagStatus')
            c = 0
            while not table.EndOfTable and c < 10:
                row = table.GetNextRow()
                vals = row.GetValues()
                status_name = {0: "None", 1: "Flagged", 2: "Complete"}
                print('  {} | FlagStatus={} ({}) | subj={}'.format(
                    store.DisplayName, vals[1], status_name.get(vals[1], '?'), vals[0]))
                c += 1
    except Exception as e:
        print('{}: error {}'.format(store.DisplayName, e))
