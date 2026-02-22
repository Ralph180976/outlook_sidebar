# -*- coding: utf-8 -*-
"""Checks GitHub Releases for newer versions of InboxBar."""

import threading
import json
from sidebar.core.config import VERSION

GITHUB_REPO = "Ralph180976/outlook_sidebar"
GITHUB_API_URL = "https://api.github.com/repos/{}/releases/latest".format(GITHUB_REPO)
RELEASES_PAGE = "https://github.com/{}/releases/latest".format(GITHUB_REPO)


def _parse_version(v_str):
    """Parse 'v1.3.18' into (1, 3, 18) tuple for comparison."""
    try:
        clean = v_str.strip().lstrip("v")
        parts = clean.split(".")
        return tuple(int(p) for p in parts)
    except:
        return (0, 0, 0)


def check_for_update(callback):
    """Check GitHub for newer version. Runs in background thread.
    
    Calls callback(latest_version, download_url, release_page) if update available.
    Calls callback(None, None, None) if up to date or check fails.
    """
    def _check():
        try:
            import urllib.request
            req = urllib.request.Request(
                GITHUB_API_URL,
                headers={"Accept": "application/vnd.github.v3+json", "User-Agent": "InboxBar"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            
            latest_tag = data.get("tag_name", "")
            current = _parse_version(VERSION)
            latest = _parse_version(latest_tag)
            
            if latest > current:
                # Find the .exe installer asset
                download_url = ""
                for asset in data.get("assets", []):
                    if asset["name"].endswith(".exe"):
                        download_url = asset["browser_download_url"]
                        break
                
                # Fallback to release page if no .exe found
                if not download_url:
                    download_url = data.get("html_url", RELEASES_PAGE)
                
                callback(latest_tag, download_url, data.get("html_url", RELEASES_PAGE))
            else:
                callback(None, None, None)
                
        except Exception as e:
            print("Update check failed: {}".format(e))
            callback(None, None, None)
    
    t = threading.Thread(target=_check, daemon=True)
    t.start()
