import win32com.client
import re

def debug_flags():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) # olFolderInbox
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        
        print(f"Total items in Inbox: {items.Count}")
        print("-" * 50)
        
        count = 0
        for item in items:
            if count >= 30: break
            
            subject = getattr(item, "Subject", "[No Subject]")
            unread = getattr(item, "UnRead", False)
            
            # Check various pinning/flagging properties
            is_marked_task = getattr(item, "IsMarkedAsTask", "N/A")
            flag_status = getattr(item, "FlagStatus", "N/A") # 2 = olFlagMarked
            importance = getattr(item, "Importance", "N/A")
            
            # Some emails might be "Pinned" (Top of list) which is different from flagged
            # But "Flagged" usually refers to FlagStatus
            
            is_flagged = (flag_status == 2)
            
            if is_flagged or is_marked_task == True:
                print(f"Subject: {subject}")
                print(f"  UnRead: {unread}")
                print(f"  IsMarkedAsTask: {is_marked_task}")
                print(f"  FlagStatus: {flag_status}")
                print(f"  Importance: {importance}")
                print("-" * 30)
                count += 1
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug_flags()
