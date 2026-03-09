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
py -3 -m PyInstaller InboxBar.spec --noconfirm --clean
```
The `--clean` flag removes cached build files to prevent stale bytecode issues.
Wait for "Build complete!" message.

## 4. Test the built exe BEFORE packaging
```bash
dist\InboxBar\InboxBar.exe
```
**CRITICAL**: Verify the exe shows emails (dark theme, Outlook + Graph emails) before continuing.
Kill the test exe before proceeding.

## 5. Build installer with Inno Setup
```bash
& "C:\Users\Ralph.SOUTHERNC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" setup.iss
```
This creates `installer_output\InboxBar_Setup.exe`.

## 6. Commit and push
```bash
git add -A
git commit -m "v1.X.Y - Brief changelog description"
git push
```

## 7. Create GitHub Release with installer attached
```bash
gh release create v1.X.Y installer_output/InboxBar_Setup.exe --title "v1.X.Y" --notes "Changelog description"
```
This uploads the installer to GitHub Releases. Existing users see an in-app update notification on next launch.

## Important Notes
- **Both version files must match**: `sidebar/core/config.py` (with `v` prefix) and `setup.iss` (without prefix)
- **Inno Setup location**: `C:\Users\Ralph.SOUTHERNC\AppData\Local\Programs\Inno Setup 6\ISCC.exe`
- **GitHub CLI**: `gh` is installed and authenticated
- **Update checker**: The app checks `api.github.com/repos/Ralph180976/outlook_sidebar/releases/latest` on startup
- **Repo is PUBLIC** — the GitHub API and Releases page work without authentication
- Version format: `v1.X.Y` (increment Y for minor updates, X for larger changes)
- Ensure `$env:Path` includes gh: `$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")`

## Lessons Learned (Do NOT Repeat)
- **Always use `--clean` flag** with PyInstaller to avoid stale build cache
- **Always test the built exe** before packaging with Inno Setup or committing — source working ≠ exe working
- **Always check with the user** before committing, version-bumping, and releasing
- **Keep `InboxBar.spec` hiddenimports in sync** — if you add a new Python module under `sidebar/`, add it to `hiddenimports` in the spec file
- **Outlook Table API limitations**: `Body` and `PR_PREVIEW` (`0x3FD9001F`) CANNOT be used as Table columns — they cause `GetValues()` to fail silently for every row, breaking all email loading
- **Installer filename is stable**: `InboxBar_Setup.exe` (no version in filename) since the download URL must stay constant for in-app updates
