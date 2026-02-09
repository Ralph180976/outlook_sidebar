import json
import io

try:
    with io.open('config.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    changed = False
    for btn in data.get('buttons', []):
        act1 = btn.get('action1', '')
        act2 = btn.get('action2', '')
        label = btn.get('label', '')
        
        # Smart Map based on action/label
        if "Delete" in act1 or "Delete" in act2 or "Trash" in label or "Delete" in label:
            btn['icon'] = u"\u2713" # Checkmark (or X?) Original was X (u2715) or Check? Code said "✕" which is u2715
            changed = True
        elif "Reply" in act1 or "Reply" in label:
            btn['icon'] = u"\u21a9" # Arrow Left
            changed = True
        elif "Open" in act1 or "Open" in label:
            btn['icon'] = u"\u2197" # Arrow Up Right
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
        with io.open('config.json', 'w', encoding='utf-8') as f:
            s = json.dumps(data, indent=4, ensure_ascii=False)
            f.write(s)
        print("Config migrated successfully.")
    else:
        print("No migration needed.")

except Exception as e:
    print("Migration failed: {}".format(e))
