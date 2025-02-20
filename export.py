import os
import shutil
import tkinter as tk
from tkinter import filedialog

# Define Paths
base_dir = os.path.expanduser("~/PDF_Extractor")
exports_dir = os.path.join(base_dir, "Exports")

def export_pdfs():
    """Prompts user to select a folder and moves extracted PDFs."""
    if not os.path.exists(exports_dir) or not os.listdir(exports_dir):
        print("⚠ No extracted PDFs found. Export skipped.")
        return

    # Open folder selection dialog
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window
    destination_folder = filedialog.askdirectory(title="Select Destination Folder for Export")

    if not destination_folder:
        print("❌ Export canceled by user.")
        return

    print(f"📂 Exporting files to: {destination_folder}")

    for file_name in os.listdir(exports_dir):
        src_path = os.path.join(exports_dir, file_name)
        dest_path = os.path.join(destination_folder, file_name)

        if os.path.exists(dest_path):
            base, ext = os.path.splitext(file_name)
            dest_path = os.path.join(destination_folder, f"{base}_Copy{ext}")
            print(f"⚠ File {file_name} already exists. Renaming to {os.path.basename(dest_path)}")

        try:
            shutil.move(src_path, dest_path)
            print(f"✅ Moved: {file_name} → {dest_path}")
        except Exception as e:
            print(f"❌ Error moving {file_name}: {e}")

    print("🚀 Export completed successfully!")

if __name__ == "__main__":
    export_pdfs()
