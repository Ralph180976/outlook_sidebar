import sys
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
print("Fetching 365 task...")
tasks = client.get_tasks(due_filters=["Overdue", "Today", "Tomorrow", "No Date"])
gtask = next((t for t in tasks if t.get('entry_id') and len(t['entry_id']) > 40), None)
if gtask:
    print("Found Graph Task:", gtask['subject'])
    print("Marking complete...")
    res = client.mark_task_complete(gtask['entry_id'])
    print("Result:", res)
else:
    print("No Graph Task found.")

print("\nFetching 365 flagged email...")
flags, _ = client.graph.get_inbox_items(count=20, unread_only=False, only_flagged=True, due_filters=["No Date"])
if flags:
    print("Found Graph Flagged Email:", flags[0]['subject'])
    print("Unflagging...")
    res = client.unflag_email(flags[0]['entry_id'])
    print("Result:", res)
else:
    print("No Graph flagged email found.")
