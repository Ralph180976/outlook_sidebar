import sys
from sidebar.services.hybrid_client import HybridMailClient
client = HybridMailClient()
print("Getting emails...")
emails, unread = client.get_inbox_items(count=3)
print(f"Got {len(emails)} emails")

entry_id = None
store_id = None

for e in emails:
    # Just checking the entry ID
    if len(e["entry_id"]) > 100:  # Likely graph
        print("Found a graph email:", e["subject"])
        entry_id = e["entry_id"]
        store_id = e.get("store_id")
        break

if entry_id:
    print("Testing toggle_flag on Graph API email...")
    res = client.toggle_flag(entry_id, store_id)
    print("Result:", res)
    
    print("Testing reply_to_email on Graph API email...")
    res2 = client.reply_to_email(entry_id, store_id)
    print("Result:", res2)
else:
    print("No Graph API emails found.")
