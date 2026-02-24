import sys
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
print("Trying to find a webLink in ToDo response...")
tasks = client.graph._request("GET", "/me/todo/lists")
list_id = tasks["value"][0]["id"]
ts = client.graph._request("GET", f"/me/todo/lists/{list_id}/tasks")
print(ts["value"][0].keys())

print("Trying old outlook tasks API instead...")
try:
    old_tasks = client.graph._request("GET", "/me/outlook/tasks?$top=1")
    if old_tasks and "value" in old_tasks and len(old_tasks["value"]) > 0:
         print(old_tasks["value"][0].keys())
         print("Old WebLink:", old_tasks["value"][0].get("webLink"))
except Exception as e:
    print("Old API error:", e)

