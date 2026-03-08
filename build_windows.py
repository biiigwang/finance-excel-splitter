#!/usr/bin/env python3
"""
Build script for Windows GUI app.

This script packages the GUI application as a standalone Windows .exe file.
Run this on Windows: python build_windows.py

Or use the generated .bat file: build_windows.bat
"""

import subprocess
import sys
import shutil
from pathlib import Path


def main():
    """Build the Windows executable."""
    print("=" * 60)
    print("Building Windows App - 财务数据拆分工具")
    print("=" * 60)

    # Detect Python version for naming
    py_version = f"{sys.version_info.major}.{sys.version_info.minor}"
    if sys.version_info >= (3, 10):
        suffix = "-win10+"
    else:
        suffix = "-win7"

    app_name = f"财务数据拆分工具{suffix}"

    # Clean previous builds
    print("\nCleaning previous builds...")
    for dir_name in ['build', 'dist']:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"  Removed {dir_name}/")

    # Check if pyinstaller is installed
    try:
        subprocess.run(['pyinstaller', '--version'], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("\nPyInstaller not found. Installing...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)

    # PyInstaller command
    print(f"\nRunning PyInstaller (Python {py_version})...")
    cmd = [
        'pyinstaller',
        '--name', app_name,
        '--windowed',
        '--onefile',
        '--clean',
        '--noconfirm',
        '--icon=icons/app.ico',
        'gui_app.py'
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"\nBuild failed:\n{result.stderr}")
        sys.exit(1)

    print("Build successful!")

    # Check output
    exe_path = Path(f'dist/{app_name}.exe')
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / 1024 / 1024
        print(f"\nExecutable created: {exe_path}")
        print(f"Size: {size_mb:.1f} MB")
        print(f"\nOutput location: {exe_path.absolute()}")
    else:
        # Try default name
        exe_path = Path('dist/财务数据拆分工具.exe')
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / 1024 / 1024
            print(f"\nExecutable created: {exe_path}")
            print(f"Size: {size_mb:.1f} MB")
        else:
            print("\nWarning: Could not find .exe file in dist/")
            print("Contents of dist/:", list(Path('dist').glob('*')))

    print("\n" + "=" * 60)
    print("Build complete!")
    print("=" * 60)

    return 0


if __name__ == '__main__':
    sys.exit(main())