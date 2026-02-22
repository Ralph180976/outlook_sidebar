---
description: Update version, build installer, and push to GitHub
---

# Version Update, Build & Deploy Workflow

// turbo-all

Follow these steps when incrementing the version:

## 1. Update VERSION constant in config
Edit `sidebar/core/config.py` line 22:
```python
VERSION = "v1.X.Y" # Brief changelog description
```

## 2. Update VERSION in Inno Setup script
Edit `setup.iss` line 5:
```
#define MyAppVersion "1.X.Y"
```
Make sure both versions match (config.py has the `v` prefix, setup.iss does not).

## 3. Build with PyInstaller
```bash
py -3 -m PyInstaller InboxBar.spec --noconfirm
```
Wait for "Build complete!" message.

## 4. Build installer with Inno Setup
```bash
& "C:\Users\Ralph.SOUTHERNC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" setup.iss
```
This creates `installer_output\InboxBar_Setup_v1.X.Y.exe` — the proper Windows installer that registers in Add & Remove Programs.

## 5. Commit and push
```bash
git add -A
git commit -m "v1.X.Y - Brief changelog description"
git push
```

## 6. Create GitHub Release with installer attached
```bash
gh release create v1.X.Y installer_output/InboxBar_Setup_v1.X.Y.exe --title "v1.X.Y" --notes "Changelog description"
```
This uploads the installer to GitHub Releases. Existing users will see an in-app update notification on next launch.

## Important Notes
- **Both version files must match**: `sidebar/core/config.py` (with `v` prefix) and `setup.iss` (without prefix)
- **Inno Setup location**: `C:\Users\Ralph.SOUTHERNC\AppData\Local\Programs\Inno Setup 6\ISCC.exe`
- **GitHub CLI**: `gh` is installed via winget. If auth is needed, run `gh auth login`
- **Update checker**: The app checks `api.github.com/repos/Ralph180976/outlook_sidebar/releases/latest` on startup
- **Users see**: A notification bar above the footer when a new version is available
- Version format: `v1.X.Y` (increment Y for minor updates, X for larger changes)
