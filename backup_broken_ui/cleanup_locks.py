
import os
import signal
import sys

# Constants
LOCK_FILE = "sidebar.lock"

def cleanup():
    print("--- CLEANUP START ---")
    
    # 1. Check Lock File
    if os.path.exists(LOCK_FILE):
        print("Found lock file: {}".format(LOCK_FILE))
        try:
            with open(LOCK_FILE, "r") as f:
                content = f.read().strip()
                if content.isdigit():
                    pid = int(content)
                    print("clean_locks: Found PID in lock file: {}".format(pid))
                    
                    # Try to kill it
                    try:
                        os.kill(pid, signal.SIGTERM)
                        print("clean_locks: Sent SIGTERM to PID {}".format(pid))
                    except OSError as e:
                        print("clean_locks: Could not kill PID {} (likely not running): {}".format(pid, e))
                else:
                    print("clean_locks: Lock file content invalid: '{}'".format(content))
                    
            # Remove the file
            os.remove(LOCK_FILE)
            print("clean_locks: Removed {}".format(LOCK_FILE))
            
        except Exception as e:
            print("clean_locks: Error processing lock file: {}".format(e))
    else:
        print("No lock file found.")

    print("--- CLEANUP END ---")

if __name__ == "__main__":
    cleanup()
