import sys
from sidebar.services.graph_auth import GraphAuth

auth = GraphAuth()
accts = auth.app.get_accounts()
print(f"Found {len(accts)} accounts in MSAL cache:")
for a in accts:
    print(f"- {a.get('username')}, {a.get('name')}")
