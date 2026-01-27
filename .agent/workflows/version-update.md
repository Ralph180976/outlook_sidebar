---
description: Update version and push to GitHub
---

# Version Update and GitHub Push Workflow

Follow these steps when incrementing the version:

## 1. Update VERSION constant
Edit `sidebar_main.py` line 29:
```python
VERSION = "v1.0.X"  # Increment the version number
```

## 2. Restart the sidebar
// turbo
Run: `py -3 sidebar_main.py`

This ensures the new version displays in the footer.

## 3. Commit changes
```bash
git add sidebar_main.py
git commit -m "vX.X.X - [Brief description]"
```

## 4. Push to GitHub
```bash
git push
```

## Important Notes
- **Always restart the sidebar** after updating VERSION to verify the footer shows the new version
- The VERSION constant (line 29) controls what's displayed in the footer
- Version format: `v1.0.X` (increment the last number for minor updates)
