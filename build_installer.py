"""
InboxBar Build Script
=====================
Automates the full build pipeline:
1. Build executable with PyInstaller
2. Create installer with Inno Setup (if available)

Usage:
    py -3 build_installer.py
"""

import os
import sys
import shutil
import subprocess
import time

# --- Configuration ---
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
SPEC_FILE = os.path.join(PROJECT_DIR, "InboxBar.spec")
DIST_DIR = os.path.join(PROJECT_DIR, "dist")
BUILD_DIR = os.path.join(PROJECT_DIR, "build")
INSTALLER_OUTPUT = os.path.join(PROJECT_DIR, "installer_output")
SETUP_ISS = os.path.join(PROJECT_DIR, "setup.iss")

# Inno Setup paths (common install locations)
INNO_PATHS = [
    os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Inno Setup 6", "ISCC.exe"),
    os.path.join(os.environ.get("ProgramFiles", ""), "Inno Setup 6", "ISCC.exe"),
    os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Inno Setup 5", "ISCC.exe"),
]


def find_inno_setup():
    """Find Inno Setup compiler."""
    # Check PATH first
    result = shutil.which("iscc") or shutil.which("ISCC")
    if result:
        return result
    
    # Check common install locations
    for path in INNO_PATHS:
        if os.path.exists(path):
            return path
    
    return None


def kill_running():
    """Kill any running InboxBar instances."""
    print("\n[1/4] Killing running instances...")
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "InboxBar.exe"],
            capture_output=True, timeout=5
        )
        print("      Killed running InboxBar")
    except Exception:
        print("      No running instances found")


def build_pyinstaller():
    """Build the executable with PyInstaller."""
    print("\n[2/4] Building executable with PyInstaller...")
    print("      This may take a few minutes...\n")
    
    # Clean old build
    for d in [BUILD_DIR, os.path.join(DIST_DIR, "InboxBar")]:
        if os.path.exists(d):
            print(f"      Cleaning {d}...")
            shutil.rmtree(d, ignore_errors=True)
    
    # Run PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        SPEC_FILE
    ]
    
    result = subprocess.run(cmd, cwd=PROJECT_DIR)
    
    if result.returncode != 0:
        print("\n[ERROR] PyInstaller build failed!")
        return False
    
    # Verify output
    exe_path = os.path.join(DIST_DIR, "InboxBar", "InboxBar.exe")
    if not os.path.exists(exe_path):
        print(f"\n[ERROR] Expected output not found: {exe_path}")
        return False
    
    # Get size
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(os.path.join(DIST_DIR, "InboxBar")):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            total_size += os.path.getsize(fp)
    
    print(f"\n      Build successful!")
    print(f"      Output: {exe_path}")
    print(f"      Total size: {total_size / (1024*1024):.1f} MB")
    return True


def build_installer():
    """Build the installer with Inno Setup."""
    iscc = find_inno_setup()
    
    if not iscc:
        print("\n[3/4] Inno Setup not found - skipping installer creation")
        print("      To create a proper installer, install Inno Setup 6 from:")
        print("      https://jrsoftware.org/isdl.php")
        print("\n      For now, the portable build is available at:")
        print(f"      {os.path.join(DIST_DIR, 'InboxBar')}")
        return False
    
    print(f"\n[3/4] Building installer with Inno Setup...")
    print(f"      Compiler: {iscc}")
    
    # Create output directory
    os.makedirs(INSTALLER_OUTPUT, exist_ok=True)
    
    # Run Inno Setup compiler
    result = subprocess.run([iscc, SETUP_ISS], cwd=PROJECT_DIR)
    
    if result.returncode != 0:
        print("\n[ERROR] Inno Setup compilation failed!")
        return False
    
    # Find the output
    for f in os.listdir(INSTALLER_OUTPUT):
        if f.endswith(".exe"):
            full_path = os.path.join(INSTALLER_OUTPUT, f)
            size_mb = os.path.getsize(full_path) / (1024 * 1024)
            print(f"\n      Installer created successfully!")
            print(f"      {full_path}")
            print(f"      Size: {size_mb:.1f} MB")
            return True
    
    return False


def create_portable_zip():
    """Create a portable ZIP package as a fallback."""
    print("\n[4/4] Creating portable ZIP package...")
    
    source = os.path.join(DIST_DIR, "InboxBar")
    if not os.path.exists(source):
        print("      Skipping - no build output found")
        return False
    
    os.makedirs(INSTALLER_OUTPUT, exist_ok=True)
    zip_path = os.path.join(INSTALLER_OUTPUT, "InboxBar_Portable_v1.3.15")
    
    shutil.make_archive(zip_path, 'zip', DIST_DIR, 'InboxBar')
    
    size_mb = os.path.getsize(zip_path + ".zip") / (1024 * 1024)
    print(f"      Portable ZIP: {zip_path}.zip")
    print(f"      Size: {size_mb:.1f} MB")
    return True


def main():
    print("=" * 60)
    print("  InboxBar Build Pipeline")
    print("=" * 60)
    
    start = time.time()
    
    # Step 1: Kill running instances
    kill_running()
    
    # Step 2: Build with PyInstaller
    if not build_pyinstaller():
        print("\nBuild failed. Aborting.")
        sys.exit(1)
    
    # Step 3: Build installer (if Inno Setup available)
    installer_built = build_installer()
    
    # Step 4: Create portable ZIP as fallback
    create_portable_zip()
    
    elapsed = time.time() - start
    
    print("\n" + "=" * 60)
    print(f"  Build complete in {elapsed:.0f}s")
    print("=" * 60)
    
    if installer_built:
        print(f"\n  Installer: {INSTALLER_OUTPUT}")
    else:
        print(f"\n  Portable build: {os.path.join(DIST_DIR, 'InboxBar')}")
        print(f"  Portable ZIP:   {INSTALLER_OUTPUT}")
        print("\n  To create a proper installer, install Inno Setup 6:")
        print("  https://jrsoftware.org/isdl.php")
        print("  Then run this script again.")
    
    print()


if __name__ == "__main__":
    main()
