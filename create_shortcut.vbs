Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = "C:\Users\Ralph.SOUTHERNC\Desktop\InboxBar.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "C:\Users\Ralph.SOUTHERNC\AppData\Local\Programs\Python\Python312\pythonw.exe"
oLink.Arguments = """c:\Dev\Outlook_Sidebar\sidebar_main.py"""
oLink.WorkingDirectory = "c:\Dev\Outlook_Sidebar"
oLink.IconLocation = "c:\Dev\Outlook_Sidebar\icons\inboxbar.ico,0"
oLink.Description = "InboxBar"
oLink.Save
