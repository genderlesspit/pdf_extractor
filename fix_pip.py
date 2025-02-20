# fix_pip.py

import os
import subprocess
import sys
import pip

def fix_pip_vendor():
    vendor_dir = os.path.join(os.path.dirname(pip.__file__), "_vendor")
    if vendor_dir not in sys.path:
        sys.path.insert(0, vendor_dir)
        print(f"Added vendor directory to sys.path: {vendor_dir}")

    rich_dir = os.path.join(vendor_dir, "rich")
    markdown_path = os.path.join(rich_dir, "markdown.py")
    
    if not os.path.exists(markdown_path):
        print("rich.markdown not found. Reinstalling 'rich' from source...")
        # Force a source installation of rich
        subprocess.check_call([
            sys.executable,
            "-m", "pip",
            "install",
            "--force-reinstall",
            "--no-binary", ":all:",
            "rich"
        ])
    else:
        print("rich.markdown is present.")

if __name__ == "__main__":
    try:
        fix_pip_vendor()
        print("pip vendor fix completed successfully.")
    except Exception as e:
        print(f"Error fixing pip vendor: {e}")
        sys.exit(1)
