#!/usr/bin/env python3
"""
Build script for macOS GUI app.

This script packages the GUI application as a standalone macOS .app bundle.
"""

import subprocess
import sys
import shutil
from pathlib import Path


def main():
    """Build the macOS application."""
    print("=" * 60)
    print("Building macOS App - 财务数据拆分工具")
    print("=" * 60)

    # Clean previous builds
    print("\nCleaning previous builds...")
    for dir_name in ['build', 'dist']:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"  Removed {dir_name}/")

    # PyInstaller command
    print("\nRunning PyInstaller...")
    cmd = [
        'pyinstaller',
        '--name=财务数据拆分工具',
        '--windowed',
        '--onefile',
        '--clean',
        '--noconfirm',
        'gui_app.py'
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"\nBuild failed:\n{result.stderr}")
        sys.exit(1)

    print("Build successful!")

    # Check output
    app_path = Path('dist/财务数据拆分工具.app')
    if app_path.exists():
        print(f"\nApp bundle created: {app_path}")
        print(f"Size: {app_path.stat().st_size / 1024 / 1024:.1f} MB")
    else:
        exe_path = Path('dist/财务数据拆分工具')
        if exe_path.exists():
            print(f"\nExecutable created: {exe_path}")
            print(f"Size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")

    print("\n" + "=" * 60)
    print("Build complete!")
    print("=" * 60)


if __name__ == '__main__':
    main()
