import tkinter as tk
from tkinter import filedialog
import shutil
import os

# Define Paths
base_dir = os.path.expanduser("~/PDF_Extractor/Master_PDFs")
os.makedirs(base_dir, exist_ok=True)

# Function to open file dialog and copy selected PDF
def upload_pdf():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file_path:
        print("❌ No file selected.")
        return
    
    file_name = os.path.basename(file_path)
    dest_path = os.path.join(base_dir, file_name)

    try:
        shutil.copy(file_path, dest_path)
        print(f"✅ Successfully uploaded: {file_name} → {dest_path}")
    except Exception as e:
        print(f"❌ Error copying file: {e}")

if __name__ == "__main__":
    upload_pdf()
