import sys
from sidebar.services.hybrid_client import HybridMailClient
import time

client = HybridMailClient()
print("Re-fetching lists directly avoiding cache...")
try:
    lists = client.graph._request("GET", "/me/todo/lists")
    for lst in lists["value"]:
        print(f"List: {lst['displayName']}, ID: {lst['id']}")
        tasks = client.graph._request("GET", f"/me/todo/lists/{lst['id']}/tasks")
        if tasks and "value" in tasks:
            for t in tasks["value"]:
                 print(f"    - Task: {t.get('title')}, Status: {t.get('status')}, Due: {t.get('dueDateTime')}")
except Exception as e:
    print(e)
