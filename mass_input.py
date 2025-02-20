import tkinter as tk
from tkinter import simpledialog
import os
import csv
import sys

# Define Paths
base_dir = os.path.expanduser("~/PDF_Extractor")
grid_update_file = os.path.join(base_dir, "temp_grid_update.csv")  # File for PowerShell to read

# Ensure base directory exists
os.makedirs(base_dir, exist_ok=True)

# Function to handle mass input for different data types
def mass_input(input_type):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    prompt_messages = {
        "master_pdf": "Enter Master PDFs (one per line):",
        "output_name": "Enter Output Names (one per line):",
        "page_range": "Enter Page Ranges (one per line):"
    }

    # Determine prompt based on input type
    input_text = simpledialog.askstring("Mass Input", prompt_messages.get(input_type, "Enter values (one per line):"))

    if not input_text:
        print("❌ No input provided.")
        return

    values = [val.strip() for val in input_text.split("\n") if val.strip()]

    # Load existing grid data if available
    existing_data = []
    if os.path.exists(grid_update_file):
        with open(grid_update_file, "r", newline="", encoding="utf-8") as file:
            reader = csv.reader(file)
            existing_data = list(reader)

    # Determine the column index to update
    column_index = {"master_pdf": 0, "output_name": 1, "page_range": 2}.get(input_type, 0)

    # Update the correct column while keeping other columns intact
    for i, value in enumerate(values):
        if i < len(existing_data):
            existing_data[i][column_index] = value
        else:
            row = ["", "", ""]
            row[column_index] = value
            existing_data.append(row)

    # Save updated data to the CSV
    with open(grid_update_file, "w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(existing_data)

    print(f"✅ Successfully updated {len(values)} entries for {input_type}.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        mass_input(sys.argv[1])
    else:
        print("❌ No input type specified.")
