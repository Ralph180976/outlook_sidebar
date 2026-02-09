import json
import os
import io

icons_map = {
    "Reply": "Reply.png",
    "Open Email": "open.png",
    "Delete": "Delete.png",
    "Flag": "Flag.png",
    "Move To...": "Move to Folder.png",
    "Mark Read": "Mark as Read.png",
    u"\u2713": "Delete.png", # Checkmark -> Delete (or Mark Read?)
    u"\u2714": "Delete.png", # Heavy Checkmark
    u"\u2610": "Mark as Read.png", # Ballot Box
    u"\u2709": "open.png", # Envelope
    u"\u21a9": "Reply.png", # Arrow
    u"\u2197": "open.png" # Arrow Up Right
}

# Combo actions
combo_map = {
    ("Mark Read", "Delete"): "Read & Delete.png"
}

try:
    if not os.path.exists("config.json"):
        print("Config not found.")
        exit()

    with io.open('config.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    changed = False
    for btn in data.get('buttons', []):
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
        with io.open('config.json', 'w', encoding='utf-8') as f:
            # json.dump in Py2 writes bytes if ensure_ascii=False, but io.open expects unicode.
            # standard json.dump writes ascii by default which is safe.
            # But if we want pretty indentation...
            # Construct string first
            s = json.dumps(data, indent=4, ensure_ascii=False)
            f.write(s)
        print("PNG Migration successful.")
    else:
        print("No changes needed.")

except Exception as e:
    print("Migration error: {}".format(e))
