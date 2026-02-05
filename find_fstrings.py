
import re

print("--- Scanning for f-strings ---")
try:
    with open("sidebar_main.py", "r") as f:
        lines = f.readlines()

    count = 0
    for i, line in enumerate(lines):
        # Look for f" or f' (but careful with comments, though simple check is fine for now)
        # We need to match f followed primarily by quote
        if re.search(r'[^a-zA-Z0-9]f["\']|^f["\']', line):
             print("Line {}: {}".format(i+1, line.strip()))
             count += 1
    
    print("--- Found {} potential f-strings ---".format(count))

except Exception as e:
    print("Error: {}".format(e))
