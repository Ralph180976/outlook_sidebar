import sys
from sidebar.services.graph_auth import GraphAuth

auth = GraphAuth()
accts = auth.get_accounts()
if accts:
    print("Authority type:", accts[0].get('authority_type'))
    print("Tenant profile:", accts[0].get('tenant_profiles'))
    print("Realm:", accts[0].get('realm'))
