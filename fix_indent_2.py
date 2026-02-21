
import os

target_file = r'c:\Dev\Outlook_Sidebar\sidebar_main.py'

try:
    with open(target_file, 'r') as f:
        lines = f.readlines()

    with open(target_file, 'w') as f:
        for i, line in enumerate(lines):
            ln = i + 1
            # Range 2723 to 2919 (inclusive)
            if 2723 <= ln < 2920:
                # Add one leading space
                f.write(' ' + line)
            else:
                f.write(line)
    
    print("Indentation fixed (2).")

except Exception as e:
    print("Error: {}".format(e))
