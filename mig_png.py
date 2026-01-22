import json
import os

icons_map = {
    "Reply": "Reply.png",
    "Open Email": "open.png",
    "Delete": "Delete.png",
    "Flag": "Flag.png",
    "Move To...": "Move to Folder.png",
    "Mark Read": "Mark as Read.png"
}

# Combo actions
combo_map = {
    ("Mark Read", "Delete"): "Read & Delete.png"
}

try:
    if not os.path.exists("sidebar_config.json"):
        print("Config not found.")
        exit()

    with open('sidebar_config.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    changed = False
    for btn in data.get('btn_config', []):
        act1 = btn.get('action1', 'None')
        act2 = btn.get('action2', 'None')
        
        # Check specific combo first
        if act1 == "Mark Read" and act2 == "Delete":
             if os.path.exists(os.path.join("icons", "Read & Delete.png")):
                 btn['icon'] = "Read & Delete.png"
                 changed = True
                 continue

        # Check primary action
        if act1 in icons_map:
            png = icons_map[act1]
            if os.path.exists(os.path.join("icons", png)):
                btn['icon'] = png
                changed = True

    if changed:
        with open('sidebar_config.json', 'w', encoding='utf-8') as f:
            json.dump(data, f)
        print("PNG Migration successful.")
    else:
        print("No changes needed.")

except Exception as e:
    print(f"Migration error: {e}")
