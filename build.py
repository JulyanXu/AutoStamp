"""
Build script for AutoStamp.
Run on Windows to create distributable package.
On macOS, creates a local build for testing.

Usage: python build.py
"""
import os
import platform
import subprocess
import sys


def main():
    system = platform.system()
    app_name = "AutoStamp"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", app_name,
        "--windowed",
        "--noconfirm",
        "--add-data", f"resources{os.pathsep}resources",
        "main.py",
    ]

    print(f"Building {app_name} for {system}...")
    subprocess.run(cmd, check=True)

    dist_dir = os.path.join("dist", app_name)
    print(f"\nBuild complete: {dist_dir}")

    if system == "Windows":
        print("\nNext steps:")
        print(f"1. Zip the entire '{dist_dir}' folder for distribution")
        print(f"2. Users extract the zip and run {app_name}.exe")
        print(f"   (Microsoft Word/Excel or WPS Office must be installed on the target machine for .docx/.xlsx conversion)")


if __name__ == "__main__":
    main()
