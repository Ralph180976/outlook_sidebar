#!/usr/bin/env python3
"""
commit.py - Automated version bumping and Git commit helper.

Usage:
    py -3 commit.py fix "Fixed the Outlook button issue"
    py -3 commit.py feat "Added new settings panel"
    py -3 commit.py breaking "Redesigned configuration format"

This will:
1. Parse the current VERSION from sidebar_main.py
2. Bump the appropriate version number (MAJOR.MINOR.PATCH)
3. Update sidebar_main.py with the new version
4. Stage all changes and create a commit with a formatted message
"""

import re
import subprocess
import sys
import os

VERSION_FILE = "sidebar/core/config.py"
VERSION_PATTERN = r'VERSION\s*=\s*["\']v?(\d+)\.(\d+)\.(\d+)["\']'

def get_current_version():
    """Reads the current version from sidebar_main.py"""
    with open(VERSION_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    
    match = re.search(VERSION_PATTERN, content)
    if not match:
        print("ERROR: Could not find VERSION in sidebar_main.py")
        sys.exit(1)
    
    return int(match.group(1)), int(match.group(2)), int(match.group(3))

def set_version(major, minor, patch):
    """Writes the new version to sidebar_main.py"""
    with open(VERSION_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    
    new_version = f'VERSION = "v{major}.{minor}.{patch}"'
    new_content = re.sub(VERSION_PATTERN.replace(r'v?', r'v'), new_version, content)
    
    # Handle case where 'v' prefix might be missing
    if new_content == content:
        new_content = re.sub(VERSION_PATTERN, new_version, content)
    
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(new_content)
    
    return f"v{major}.{minor}.{patch}"

def bump_version(change_type, major, minor, patch):
    """
    Bumps the version based on change type:
    - fix: Bump PATCH (1.0.1 -> 1.0.2)
    - feat: Bump MINOR, reset PATCH (1.0.2 -> 1.1.0)
    - breaking: Bump MAJOR, reset MINOR and PATCH (1.1.2 -> 2.0.0)
    """
    if change_type == "fix":
        return major, minor, patch + 1
    elif change_type == "feat":
        return major, minor + 1, 0
    elif change_type == "breaking":
        return major + 1, 0, 0
    else:
        print(f"ERROR: Unknown change type '{change_type}'")
        print("Valid types: fix, feat, breaking")
        sys.exit(1)

def run_git_command(args):
    """Runs a git command and returns the output."""
    result = subprocess.run(
        ["git"] + args,
        capture_output=True,
        text=True,
        cwd=os.path.dirname(os.path.abspath(__file__)) or "."
    )
    return result.returncode, result.stdout, result.stderr

def main():
    if len(sys.argv) < 3:
        print("Usage: py -3 commit.py <fix|feat|breaking> \"Commit message\"")
        print("")
        print("Examples:")
        print('  py -3 commit.py fix "Fixed Outlook button duplicates"')
        print('  py -3 commit.py feat "Added Calendar integration"')
        print('  py -3 commit.py breaking "New config format"')
        sys.exit(1)
    
    change_type = sys.argv[1].lower()
    message = sys.argv[2]
    
    # Get current version
    major, minor, patch = get_current_version()
    old_version = f"v{major}.{minor}.{patch}"
    print(f"Current version: {old_version}")
    
    # Calculate new version
    new_major, new_minor, new_patch = bump_version(change_type, major, minor, patch)
    new_version = set_version(new_major, new_minor, new_patch)
    print(f"New version: {new_version}")
    
    # Format commit message
    prefix_map = {
        "fix": "fix",
        "feat": "feat", 
        "breaking": "BREAKING"
    }
    commit_message = f"{new_version}: [{prefix_map[change_type]}] {message}"
    print(f"Commit message: {commit_message}")
    
    # Stage all changes
    print("\nStaging changes...")
    code, out, err = run_git_command(["add", "."])
    if code != 0:
        print(f"ERROR staging: {err}")
        sys.exit(1)
    
    # Commit
    print("Creating commit...")
    code, out, err = run_git_command(["commit", "-m", commit_message])
    if code != 0:
        print(f"ERROR committing: {err}")
        sys.exit(1)
    
    print(out)
    print(f"\n✅ Successfully committed as {new_version}")
    print("Run 'git push' to upload to GitHub.")

if __name__ == "__main__":
    main()
