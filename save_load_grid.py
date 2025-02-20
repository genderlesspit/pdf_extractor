import os
import csv
import sys
import json

# Define paths
base_dir = os.path.expanduser("~/PDF_Extractor")
grid_save_file = os.path.join(base_dir, "saved_grid.csv")

def save_grid():
    """Saves the current grid data to a CSV file."""
    try:
        # Read grid data from standard input (PowerShell will send it as JSON)
        grid_data = json.loads(sys.stdin.read())

        with open(grid_save_file, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Master PDF", "Output Name", "Page Range"])  # Column headers
            for row in grid_data:
                writer.writerow(row)

        print(f"✅ Grid successfully saved to {grid_save_file}")
    except Exception as e:
        print(f"❌ Error saving grid: {e}")

def load_grid():
    """Loads saved grid data from a CSV file and prints it as JSON for PowerShell."""
    if not os.path.exists(grid_save_file):
        print("⚠ No saved grid file found.")
        return

    try:
        with open(grid_save_file, "r", newline="", encoding="utf-8") as file:
            reader = csv.reader(file)
            next(reader)  # Skip header row
            grid_data = [row for row in reader]

        print(json.dumps(grid_data))  # Send data to PowerShell in JSON format
    except Exception as e:
        print(f"❌ Error loading grid: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "load":
        load_grid()
    else:
        save_grid()
