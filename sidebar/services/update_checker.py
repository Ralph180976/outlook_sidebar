# -*- coding: utf-8 -*-
"""Checks GitHub Releases for newer versions of InboxBar."""

import threading
import json
from sidebar.core.config import VERSION

GITHUB_API_URL = "https://api.github.com/repos/Ralph180976/outlook_sidebar/releases/latest"
RELEASES_PAGE = "https://github.com/Ralph180976/outlook_sidebar/releases/latest"


def _parse_version(v_str):
    """Parse 'v1.3.18' into (1, 3, 18) tuple for comparison."""
    try:
        clean = v_str.strip().lstrip("v")
        return tuple(int(p) for p in clean.split("."))
    except:
        return (0, 0, 0)


def check_for_update(callback):
    """Check GitHub Releases for a newer version. Runs in background thread.
    
    Calls callback(latest_version, download_url) if update available.
    Calls callback(None, None) if up to date or check fails.
    """
    def _check():
        try:
            import urllib.request
            req = urllib.request.Request(
                GITHUB_API_URL,
                headers={
                    "Accept": "application/vnd.github.v3+json",
                    "User-Agent": "InboxBar"
                }
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            
            latest_tag = data.get("tag_name", "")
            
            # Find the .exe asset download URL, fall back to releases page
            download_url = RELEASES_PAGE
            for asset in data.get("assets", []):
                if asset.get("name", "").endswith(".exe"):
                    download_url = asset["browser_download_url"]
                    break
            
            current = _parse_version(VERSION)
            latest = _parse_version(latest_tag)
            
            if latest > current:
                callback(latest_tag, download_url)
            else:
                callback(None, None)
                
        except Exception as e:
            print("Update check failed: {}".format(e))
            callback(None, None)
    
    t = threading.Thread(target=_check, daemon=True)
    t.start()
