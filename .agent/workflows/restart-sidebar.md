---
description: Stop and restart the Outlook Sidebar application
---

# Restart Outlook Sidebar

// turbo-all

## Steps

1. Stop any running instances of the sidebar:
```powershell
taskkill /IM python.exe /F 2>$null; taskkill /IM pythonw.exe /F 2>$null
```

2. Start the sidebar with Python 3:
```powershell
py -3 sidebar_main.py
```

## Notes
- Always use `py -3` on this system (not `python`) because the default Python is 2.7
- The sidebar runs from `c:\Dev\Outlook_Sidebar\sidebar_main.py`
- Version number is defined in `sidebar_main.py` at line ~27 as `VERSION = "vX.Y.Z"`
- **AUTO-RUN PERMISSION**: The AI assistant has explicit permission to run this workflow automatically at its discretion to ensure the application state is current.
