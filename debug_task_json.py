import sys
import json
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
tasks = client.graph._request("GET", "/me/todo/lists")
if tasks and "value" in tasks:
    for lst in tasks["value"]:
        print(f"--- List: {lst['displayName']} ---")
        ts = client.graph._request("GET", f"/me/todo/lists/{lst['id']}/tasks")
        if ts and "value" in ts:
            for t in ts["value"]:
                print(json.dumps(t, indent=2))
        break # just check the first list
