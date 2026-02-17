import win32com.client
from datetime import datetime, timedelta

def debug_calendar():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Restrict to next 7 days (and past 30)
    now = datetime.now()
    start = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=30)
    end = start + timedelta(days=60)
    
    # Iterate all stores
    for store in namespace.Stores:
        try:
            calendar = store.GetDefaultFolder(9)
            print(f"Checking Calendar in: {store.DisplayName}")
        except: 
            print(f"No calendar for {store.DisplayName}")
            continue
            
        print(f"--- Calendar Items in {store.DisplayName} ---")
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        item = items.GetFirst()
        count = 0
        while item and count < 20:
            try:
                start_t = item.Start
                # Filter manually for target range
                if start_t.replace(tzinfo=None) >= start and start_t.replace(tzinfo=None) < end:
                    
                    subject = item.Subject
                    status = getattr(item, "ResponseStatus", "N/A")
                    meeting_status = getattr(item, "MeetingStatus", "N/A") 
                    
                    print(f"Subject: {subject}")
                    print(f"  Start: {start_t}")
                    print(f"  ResponseStatus: {status} (0=None, 1=Org, 2=Tent, 3=Acc, 4=Dec, 5=NotResp)")
                    print(f"  MeetingStatus:  {meeting_status} (0=NonMeeting/Appt, 1=Meeting, 3=Received)")
                    print("-" * 30)
                
                count += 1
            except Exception as e:
                pass
                
            item = items.GetNext()
    
if __name__ == "__main__":
    debug_calendar()
