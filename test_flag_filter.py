import win32com.client

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

for store in ns.Stores:
    try:
        inbox = store.GetDefaultFolder(6)
        
        # Test the combined DASL filter: unread OR flagged
        restrict = '@SQL=("urn:schemas:httpmail:read" = 0 OR "http://schemas.microsoft.com/mapi/proptag/0x10900003" <> 0)'
        table = inbox.GetTable(restrict)
        table.Sort('ReceivedTime', True)
        table.Columns.RemoveAll()
        table.Columns.Add('Subject')
        table.Columns.Add('UnRead')
        table.Columns.Add('FlagStatus')
        c = 0
        print("--- Store: {} ---".format(store.DisplayName))
        while not table.EndOfTable and c < 10:
            row = table.GetNextRow()
            vals = row.GetValues()
            marker = ''
            if vals[2] != 0: marker += ' [FLAGGED]'
            if vals[1]: marker += ' [UNREAD]'
            print('  {}{}'.format(vals[0], marker))
            c += 1
        print('  Items: {}'.format(c))
    except Exception as e:
        print("Store {} error: {}".format(store.DisplayName, e))
