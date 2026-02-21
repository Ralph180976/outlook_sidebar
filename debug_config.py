
import sys
import os
import json

# Adjust path to find the module
sys.path.append(os.getcwd())

from sidebar_main import SidebarWindow

def debug_config_loading():
    print("--- Debugging Config Loading ---")
    
    # 1. Inspect Default State
    print("\n[1] Initializing Window (Defaults)...")
    try:
        app = SidebarWindow()
        app.withdraw() # Hide window
        
        print("\n[2] In-Memory Configuration (After load_config):")
        print(f"reminder_meeting_dates type: {type(app.reminder_meeting_dates)}")
        print(f"reminder_meeting_dates value: {app.reminder_meeting_dates}")
        
        print(f"reminder_due_filters type: {type(app.reminder_due_filters)}")
        print(f"reminder_due_filters value: {app.reminder_due_filters}")
        
        # Check for the typo version
        if hasattr(app, 'reminder_due_filter'):
            print(f"reminder_due_filter (singular) value: {app.reminder_due_filter}")
            
    except Exception as e:
        print(f"Error initializing app: {e}")

    # 3. Inspect JSON File
    print("\n[3] JSON File Content:")
    if os.path.exists("sidebar_config.json"):
        with open("sidebar_config.json", "r") as f:
            data = json.load(f)
            print(json.dumps(data, indent=4))
    else:
        print("sidebar_config.json not found.")

if __name__ == "__main__":
    debug_config_loading()
