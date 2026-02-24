import sys
import webbrowser
from sidebar.services.graph_auth import GraphAuth
import urllib.parse

auth = GraphAuth()
email = auth.get_current_user_email()
login_hint = f"?login_hint={urllib.parse.quote(email)}" if email else ""
task_id = "AQMkADAwATM3ZmYBLWFjNTEtOTdjYS0wMAItMDAKAEYAAANITyg22N35R4N4xB_2p1IHBwAA2RxgJY1sEEGfWKWXQD-YKAAAAgESAAAA2RxgJY1sEEGfWKWXQD-YKAAAAqTiAAAA"

test_url1 = f"https://to-do.live.com/tasks/id/{task_id}{login_hint}"
print("Executing to-do.live.com:", test_url1)

import os
os.startfile(test_url1)
