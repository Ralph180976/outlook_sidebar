import win32com.client
from datetime import datetime

def debug_simple():
    print("Debug Script Started...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"Failed to dispatch Outlook: {e}")
        return
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)
    
    print(f"Calendar: {calendar.Name}")
    count = calendar.Items.Count
    print(f"Items Count: {count}")
    
    items = calendar.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    
    # Target range
    from datetime import timedelta
    now = datetime.now()
    start = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=30)
    end = start + timedelta(days=60)
    
    # Iterate first 50 or count
    limit = min(count, 50)
    
    # Warning: Indices are 1-based in COM
    found = 0
    scanned = 0
    
    item = items.GetFirst()
    while item and scanned < 50:
        try:
            start_t = item.Start
            if start_t.tzinfo: start_t = start_t.replace(tzinfo=None)
            
            if start_t >= start and start_t < end:
                print(f"MATCH: {item.Subject} ({item.Start})")
                print(f"       RespStatus: {getattr(item, 'ResponseStatus', 'N/A')}")
                print(f"       MeetStatus: {getattr(item, 'MeetingStatus', 'N/A')}")
                found += 1
        except: pass
        
        item = items.GetNext()
        scanned += 1
        
    print(f"Scanned {scanned} items, found {found} matches.")
