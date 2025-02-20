import os
import requests

# Define Paths
base_dir = os.path.expanduser("~/PDF_Extractor/dependencies")
os.makedirs(base_dir, exist_ok=True)
file_path = os.path.join(base_dir, "gui_reference.txt")

# GitHub Raw File URL (Fetching gui.txt)
github_url = "https://raw.githubusercontent.com/genderlesspit/pdf_extractor/main/gui.txt"

# Fetch content from GitHub
try:
    response = requests.get(github_url, timeout=10)
    response.raise_for_status()  # Raise an error if request fails
    content = response.text

    # Write content to a local file
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(content)

    print(f"✅ Successfully saved GitHub file to: {file_path}")

except requests.RequestException as e:
    print(f"❌ Error fetching gui.txt from GitHub: {e}")
