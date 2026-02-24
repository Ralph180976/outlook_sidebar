import string

from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
emails, _ = client.get_inbox_items(count=20)
def is_hex(s): return all(c in string.hexdigits for c in s) if hasattr(s, 'isalnum') else False

g_email = next((e for e in emails if e.get("entry_id") and (not is_hex(e["entry_id"]) or len(e["entry_id"]) <= 40)), None)

if g_email:
    print(f"Graph email: {g_email['subject']}, ID: {g_email['entry_id'][:20]}...")
    client.mark_as_read(g_email['entry_id'], g_email.get('store_id'))
    client.reply_to_email(g_email['entry_id'], g_email.get('store_id'))
    print("Done")
else:
    print("No graph email")
