import win32com.client
from datetime import datetime

app = win32com.client.Dispatch('Outlook.Application')
ns = app.GetNamespace('MAPI')

store = ns.Stores['ralph.marsh@coveya.co.uk']
inbox = store.GetDefaultFolder(6)

items = inbox.Items
items.Sort("[ReceivedTime]", True)
flagged = items.Restrict("[FlagStatus] <> 0")

print("Found {} flagged items\n".format(flagged.Count))

for i in range(min(flagged.Count, 20)):
    item = flagged.Item(i + 1)
    
    print("--- {} ---".format(item.Subject[:60]))
    print("  FlagStatus:    {} ({})".format(
        item.FlagStatus,
        {0: "No Flag", 1: "Marked/Follow-up", 2: "Complete"}.get(item.FlagStatus, "?")))
    print("  FlagIcon:      {} ({})".format(
        item.FlagIcon,
        {0: "None", 1: "Purple", 2: "Orange", 3: "Green", 4: "Yellow", 5: "Blue", 6: "Red"}.get(item.FlagIcon, "?")))
    print("  IsMarkedAsTask: {}".format(item.IsMarkedAsTask))
    
    try:
        print("  FlagRequest:   '{}'".format(item.FlagRequest))
    except: pass
    
    try:
        due = item.TaskDueDate
        if due and due.year < 3000:
            print("  TaskDueDate:   {}".format(due.strftime('%Y-%m-%d')))
        else:
            print("  TaskDueDate:   (no date)")
    except: 
        print("  TaskDueDate:   (not set)")
    
    try:
        start = item.TaskStartDate
        if start and start.year < 3000:
            print("  TaskStartDate: {}".format(start.strftime('%Y-%m-%d')))
    except: pass
    
    try:
        print("  TaskComplete:  {}%".format(item.PercentComplete))
    except: pass
    
    print()
