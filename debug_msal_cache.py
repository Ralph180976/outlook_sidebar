import sys
import json
from sidebar.services.graph_auth import GraphAuth

auth = GraphAuth()
with open(auth._cache_path, "r") as f:
    cache = json.load(f)

for k, v in cache.get("Account", {}).items():
    print(v.get("realm"), v.get("local_account_id"), v.get("home_account_id"))
