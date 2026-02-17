import win32com.client
from datetime import datetime, timedelta
import locale

def debug_locale_format():
    locale.setlocale(locale.LC_ALL, '') # Set to user's default locale
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    cal = namespace.GetDefaultFolder(9) 
    
    # 1. Create Test Appointment
    try:
        appt = outlook.CreateItem(1)
        appt.Subject = "Test Locale"
        now = datetime.now()
        start = now + timedelta(days=1)
        start = start.replace(hour=15, minute=0, second=0, microsecond=0)
        end = start + timedelta(hours=1)
        appt.Start = start
        appt.End = end
        appt.Save()
    except: pass

    # 2. Test Locale Format
    try:
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        s = start - timedelta(hours=1)
        e = end + timedelta(hours=1)
        
        # Locale Aware Format
        # %x is date, %X is time
        # But Outlook usually wants specific format.
        # Let's try %d/%m/%Y explicitly since we know it works, 
        # and compare against %x to see if it Matches.
        
        print(f"Locale Date (%x): {s.strftime('%x')}")
        print(f"Locale Time (%X): {s.strftime('%X')}")
        
        # Test Query with explicitly constructed DD/MM/YYYY HH:MM
        # This is what worked before.
        uk_query_s = s.strftime('%d/%m/%Y %H:%M')
        uk_query_e = e.strftime('%d/%m/%Y %H:%M')
        
        query = "[Start] >= '{}' AND [Start] <= '{}'".format(uk_query_s, uk_query_e)
        print(f"Testing UK Query: {query}")
        
        res = items.Restrict(query)
        found = False
        for item in res:
            if item.Subject == "Test Locale":
                print("SUCCESS! UK format works.")
                found = True
                break
        
        if not found:
            print("FAILED: UK format did not find the item (Unexpected).")

    except Exception as e:
        print(f"Error: {e}")

    # 3. Cleanup
    try:
        items = cal.Items
        to_del = items.Restrict("[Subject] = 'Test Locale'")
        for item in to_del:
            item.Delete()
            print("Deleted Test Locale.")
    except: pass

if __name__ == "__main__":
    debug_locale_format()
