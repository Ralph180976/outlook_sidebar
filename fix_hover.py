# Fix show_hover_elements - byte-level approach
with open('sidebar_main.py', 'rb') as f:
    data = f.read()

# Find the function signature
start_marker = b'def show_hover_elements(e, lp=lbl_preview, fb=frame_buttons, h=lines):'
start_idx = data.find(start_marker)
if start_idx < 0:
    print("Could not find show_hover_elements")
    exit(1)

print("Found function at byte:", start_idx)

# Find the end: the next 'def hide_hover_elements' or 'def ' at same indent
end_marker = b'def hide_hover_elements'
end_idx = data.find(end_marker, start_idx)
if end_idx < 0:
    print("Could not find end marker")
    exit(1)

# Walk back to find the start of the line before hide_hover_elements
# We need to include the blank line between them
while end_idx > 0 and data[end_idx-1:end_idx] in (b' ', b'\r', b'\n'):
    end_idx -= 1

# Now we have the exact range of show_hover_elements
old_func = data[start_idx:end_idx]
print("Old function ({} bytes):".format(len(old_func)))
print(old_func.decode('utf-8', errors='replace'))

new_func = (
    b"def show_hover_elements(e, lp=lbl_preview, fb=frame_buttons, h=lines, eid=email.get('entry_id'), sid=email.get('store_id')):\r\n"
    b"                    # 1. Show Body Preview if enabled and not permanent\r\n"
    b"                    if self.show_hover_content and not self.email_show_body and lp:\r\n"
    b"                         # Lazy-fetch body on first hover\r\n"
    b"                         if not getattr(lp, '_body_loaded', False):\r\n"
    b"                             lp._body_loaded = True\r\n"
    b"                             try:\r\n"
    b"                                 item = self.outlook_client.get_item_by_entryid(eid, sid)\r\n"
    b"                                 if item:\r\n"
    b"                                     body_text = item.Body\r\n"
    b"                                     if body_text:\r\n"
    b"                                         body_text = body_text.strip()[:500]\r\n"
    b'                                         lp.config(state="normal")\r\n'
    b'                                         lp.delete("1.0", "end")\r\n'
    b'                                         lp.insert("1.0", body_text)\r\n'
    b'                                         lp.config(state="disabled")\r\n'
    b"                             except Exception as ex:\r\n"
    b'                                 print("Hover body fetch error: {}".format(ex))\r\n'
    b"                         if not lp.winfo_ismapped():\r\n"
    b"                              lp.config(height=h) \r\n"
    b'                              lp.pack(fill="x", padx=5, pady=(0, 2)) \r\n'
    b"                    \r\n"
    b"                    # 2. Show Buttons if enabled\r\n"
    b"                    if self.buttons_on_hover:\r\n"
    b"                         if not fb.winfo_ismapped():\r\n"
    b'                              fb.pack(fill="x", expand=True, padx=2, pady=(0, 2))'
)

data = data[:start_idx] + new_func + data[end_idx:]

with open('sidebar_main.py', 'wb') as f:
    f.write(data)

print("\nDone! Replaced show_hover_elements with lazy body fetch version.")
