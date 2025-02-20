import os
import requests
import zipfile
import shutil
import subprocess
import winshell
from win32com.client import Dispatch

# Define Paths
base_dir = os.path.expanduser("~/PDF_Extractor")
dependencies_dir = os.path.join(base_dir, "dependencies")
scripts_dir = os.path.join(dependencies_dir, "scripts")
exe_output_dir = os.path.join(base_dir, "dist")
zip_file_path = os.path.join(base_dir, "pdf_extractor_latest.zip")
gui_script_path = os.path.join(scripts_dir, "gui.ps1")

os.makedirs(scripts_dir, exist_ok=True)

# GitHub API URL for latest release
repo_api_url = "https://api.github.com/repos/genderlesspit/pdf_extractor/releases/latest"

def get_latest_release():
    """Fetches the latest release zip file URL from GitHub."""
    try:
        response = requests.get(repo_api_url, timeout=10)
        response.raise_for_status()
        data = response.json()
        for asset in data.get("assets", []):
            if asset["name"].endswith(".zip"):
                return asset["browser_download_url"]
    except requests.RequestException as e:
        print(f"❌ Error fetching latest release info: {e}")
    return None

def download_latest_release():
    """Downloads the latest release zip file."""
    zip_url = get_latest_release()
    if not zip_url:
        print("❌ No valid release zip found.")
        return False

    try:
        response = requests.get(zip_url, stream=True, timeout=20)
        response.raise_for_status()
        with open(zip_file_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        print(f"✅ Downloaded latest release: {zip_file_path}")
        return True
    except requests.RequestException as e:
        print(f"❌ Error downloading release: {e}")
        return False

def extract_release():
    """Extracts the downloaded zip file into the dependencies directory."""
    if not os.path.exists(zip_file_path):
        print("❌ Zip file not found.")
        return False

    try:
        with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
            zip_ref.extractall(scripts_dir)
        print(f"✅ Extracted release to: {scripts_dir}")
        return True
    except zipfile.BadZipFile as e:
        print(f"❌ Error extracting zip file: {e}")
        return False

# Install dependencies
def install_dependencies():
    """Ensure required dependencies are installed."""
    required_packages = ["pyinstaller", "requests", "winshell", "pywin32"]

    for package in required_packages:
        try:
            subprocess.run(["pip", "install", package], check=True)
            print(f"✅ Installed: {package}")
        except subprocess.CalledProcessError:
            print(f"❌ Failed to install: {package}")

# Create standalone executable to launch gui.ps1
def create_executable():
    """Uses PyInstaller to create a .exe that runs the PowerShell GUI script."""
    os.makedirs(exe_output_dir, exist_ok=True)

    launcher_script = os.path.join(scripts_dir, "launcher.py")
    exe_name = "PDF_Extractor"

    # Create a Python script to launch gui.ps1 via PowerShell
    with open(launcher_script, "w", encoding="utf-8") as file:
        file.write(f"""import subprocess
subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", r"{gui_script_path}"])
""")

    print("🚀 Creating standalone executable...")
    try:
        subprocess.run([
            "pyinstaller",
            "--onefile",
            "--distpath", exe_output_dir,
            "--name", exe_name,
            launcher_script
        ], check=True)
        print(f"✅ Executable created: {os.path.join(exe_output_dir, exe_name)}.exe")
    except subprocess.CalledProcessError:
        print("❌ Failed to create the executable.")

# Create a desktop shortcut
def create_shortcut():
    """Creates a desktop shortcut for the generated .exe file."""
    exe_path = os.path.join(exe_output_dir, "PDF_Extractor.exe")
    shortcut_path = os.path.join(winshell.desktop(), "PDF Extractor.lnk")

    if not os.path.exists(exe_path):
        print("❌ Executable not found. Cannot create shortcut.")
        return

    try:
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = exe_path
        shortcut.WorkingDirectory = exe_output_dir
        shortcut.Description = "PDF Extractor GUI"
        shortcut.IconLocation = exe_path  # Uses the exe icon
        shortcut.Save()
        print(f"✅ Desktop shortcut created: {shortcut_path}")
    except Exception as e:
        print(f"❌ Error creating shortcut: {e}")

# Run the full setup process
def run_setup():
    print("🔄 Starting setup process...")
    if download_latest_release() and extract_release():
        install_dependencies()
        create_executable()
        create_shortcut()
        print("🎉 Setup complete!")

if __name__ == "__main__":
    run_setup()
