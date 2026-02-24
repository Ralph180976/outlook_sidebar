import sys
from sidebar.services.graph_auth import GraphAuth
auth = GraphAuth()
print("Email:", auth.get_current_user_email())
