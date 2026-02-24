import sys
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
tasks = client.graph._request("GET", "https://graph.microsoft.com/beta/me/todo/lists")
list_id = tasks["value"][0]["id"]
ts = client.graph._request("GET", f"https://graph.microsoft.com/beta/me/todo/lists/{list_id}/tasks")
if ts and "value" in ts:
    print(ts["value"][0].keys())
    for t in ts["value"]:
        print(t.get('title'), t.get('webUrl'), t.get('webLink'))
        break
