import os
import csv
import json
import re

# Define paths
base_dir = os.path.expanduser("~/PDF_Extractor")
master_pdf_dir = os.path.join(base_dir, "Master_PDFs")
grid_update_file = os.path.join(base_dir, "temp_grid_update.csv")  # Grid input from PowerShell
error_log_file = os.path.join(base_dir, "errors.json")  # Errors will be stored here

# Page range validation regex (supports formats like "1-3,5,8-10")
PAGE_RANGE_REGEX = r"^\d+(-\d+)?(,\d+(-\d+)?)*$"

def validate_inputs():
    errors = {"missing_pdfs": [], "invalid_page_ranges": [], "empty_fields": []}

    if not os.path.exists(grid_update_file):
        print("❌ No data found. Please add input to the grid.")
        return False

    with open(grid_update_file, "r", newline="", encoding="utf-8") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if not rows:
        print("❌ The grid is empty.")
        return False

    for index, row in enumerate(rows):
        try:
            pdf_name, output_name, page_range = row
            pdf_path = os.path.join(master_pdf_dir, pdf_name)

            # Check for empty required fields
            if not pdf_name or not output_name or not page_range:
                errors["empty_fields"].append(f"Row {index+1}: Missing fields.")

            # Check if the PDF file exists
            if pdf_name and not os.path.exists(pdf_path):
                errors["missing_pdfs"].append(f"Row {index+1}: {pdf_name} not found.")

            # Validate page range format
            if page_range and not re.match(PAGE_RANGE_REGEX, page_range):
                errors["invalid_page_ranges"].append(f"Row {index+1}: Invalid page range '{page_range}'.")

        except ValueError:
            errors["empty_fields"].append(f"Row {index+1}: Incorrect format.")

    # Write errors to a JSON file for PowerShell to read
    with open(error_log_file, "w", encoding="utf-8") as error_file:
        json.dump(errors, error_file, indent=4)

    if any(errors.values()):
        print("❌ Validation failed. Errors logged to errors.json.")
        return False

    print("✅ All inputs are valid.")
    return True

if __name__ == "__main__":
    valid = validate_inputs()
    exit(0 if valid else 1)  # Return exit code for PowerShell
