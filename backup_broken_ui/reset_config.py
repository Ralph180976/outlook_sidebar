import os

local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
app_dir = os.path.join(local_app_data, "OutlookSidebar")
config_path = os.path.join(app_dir, "config.json")

if os.path.exists(config_path):
    print(f"Deleting {config_path}")
    os.remove(config_path)
else:
    print("Config file not found")
