import win32com.client
try:
    print("Trying GetActiveObject...")
    app = win32com.client.GetActiveObject("Outlook.Application")
    print("GetActiveObject Success!")
    ns = app.GetNamespace("MAPI")
    print("Stores: ", ns.Stores.Count)
except Exception as e:
    print("GetActiveObject failed:", e)

try:
    print("Trying Dispatch...")
    app = win32com.client.Dispatch("Outlook.Application")
    print("Dispatch Success!")
    ns = app.GetNamespace("MAPI")
    print("Stores: ", ns.Stores.Count)
except Exception as e:
    print("Dispatch failed:", e)
