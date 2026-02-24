import sys
import os
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
tasks = client.graph.get_tasks(due_filters=["Overdue", "Today", "Tomorrow", "No Date"])
for gtask in tasks:
    print("Graph task web_link:", gtask.get('web_link'))
    url = gtask.get('web_link')
    if url.startswith("https://to-do.office.com"):
        proto_url = url.replace("https://to-do.office.com/", "ms-todo://")
        print("Starting protocol URL:", proto_url)
        os.startfile(proto_url)
    break
