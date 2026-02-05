import json

try:
    with open('sidebar_config.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    changed = False
    for btn in data.get('btn_config', []):
        act1 = btn.get('action1', '')
        act2 = btn.get('action2', '')
        label = btn.get('label', '')
        
        # Smart Map based on action/label
        if "Delete" in act1 or "Delete" in act2 or "Trash" in label or "Delete" in label:
            btn['icon'] = "✕"
            changed = True
        elif "Reply" in act1 or "Reply" in label:
            btn['icon'] = "↩"
            changed = True
        elif "Open" in act1 or "Open" in label:
            btn['icon'] = "↗"
            changed = True
        elif "Move" in act1 or "Move" in label:
            btn['icon'] = "📂"
            changed = True
        elif "Flag" in act1 or "Flag" in label:
            btn['icon'] = "⚑"
            changed = True
        elif "Mark Read" in act1:
            if btn['icon'] not in ["✕", "🗑"]: # Don't overwrite delete buttons that also mark read
                 btn['icon'] = "✓"
                 changed = True

    if changed:
        with open('sidebar_config.json', 'w', encoding='utf-8') as f:
            json.dump(data, f)
        print("Config migrated successfully.")
    else:
        print("No migration needed.")

except Exception as e:
    print(f"Migration failed: {e}")
