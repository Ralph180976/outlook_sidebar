import json
import os

CONFIG_FILE = "C:\\Dev\\Outlook_Sidebar\\sidebar_config.json"
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "r") as f:
        data = json.load(f)
    
    if "reminder_task_dates" in data and "No Date" not in data["reminder_task_dates"]:
        data["reminder_task_dates"].append("No Date")
        
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=4)
        print("Updated sidebar_config.json to include 'No Date'.")
    else:
        print("'No Date' already inside or no config found.")
