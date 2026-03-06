import pythoncom
pythoncom.CoInitialize()
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

print("=== STORES ===")
print("Stores found:", ns.Stores.Count)
for i in range(ns.Stores.Count):
    store = ns.Stores.Item(i + 1)
    print("  Store {}: {} (type={})".format(i + 1, store.DisplayName, store.ExchangeStoreType))

print()
print("=== ACCOUNTS ===")
accounts = ns.Accounts
print("Accounts found:", accounts.Count)
for i in range(accounts.Count):
    acc = accounts.Item(i + 1)
    print("  Account {}: {} ({})".format(i + 1, acc.DisplayName, acc.SmtpAddress))

print()
print("=== GRAPH AUTH ===")
try:
    from sidebar.services.graph_auth import GraphAuth
    auth = GraphAuth()
    email = auth.get_current_user_email()
    print("Graph user:", email)
    msal_accounts = auth.get_accounts()
    print("MSAL accounts:", [a.get("username") for a in msal_accounts])
except Exception as e:
    print("Graph auth error:", e)

print()
print("=== HYBRID CLIENT ===")
try:
    from sidebar.services.hybrid_client import HybridMailClient
    hybrid = HybridMailClient()
    all_accounts = hybrid.get_accounts()
    print("Hybrid accounts:", all_accounts)
    print("COM backend:", "OK" if hybrid.com else "FAILED")
    print("Graph backend:", "OK" if hybrid.graph else "FAILED")
    if hybrid.com:
        com_accounts = hybrid.com.get_accounts()
        print("COM accounts:", com_accounts)
    if hybrid.graph:
        graph_accounts = hybrid.graph.get_accounts()
        print("Graph accounts:", graph_accounts)
except Exception as e:
    print("Hybrid error:", e)
    import traceback
    traceback.print_exc()
