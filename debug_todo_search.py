import webbrowser
import urllib.parse
from sidebar.services.graph_auth import GraphAuth

auth = GraphAuth()
email = auth.get_current_user_email()
login_hint = f"&login_hint={urllib.parse.quote(email)}" if email else ""

search_url = f"https://to-do.office.com/tasks/search/365+Testing?{login_hint}"
print("Testing:", search_url)
try:
    import os
    os.startfile(search_url)
except:
    pass
