
import os

target_file = r'c:\Dev\Outlook_Sidebar\sidebar_main.py'

try:
    with open(target_file, 'r') as f:
        lines = f.readlines()

    with open(target_file, 'w') as f:
        for i, line in enumerate(lines):
            ln = i + 1
            # Range 2485 to 2919 (inclusive)
            if 2485 <= ln < 2920:
                # remove one leading space if present
                if line.startswith(' '):
                    f.write(line[1:])
                else:
                    f.write(line)
            else:
                f.write(line)
    
    print("Indentation fixed.")

except Exception as e:
    print("Error: {}".format(e))
