import sys
from sidebar.services.hybrid_client import HybridMailClient

client = HybridMailClient()
me = client.graph._request("GET", "/me")
if me:
    print("User ID:", me.get("id"))
