import os
import csv
import json
import re
import fitz  # PyMuPDF for checking actual page counts

# Define paths
base_dir = os.path.expanduser("~/PDF_Extractor")
master_pdf_dir = os.path.join(base_dir, "Master_PDFs")
grid_update_file = os.path.join(base_dir, "temp_grid_update.csv")  # Grid input from PowerShell
error_log_file = os.path.join(base_dir, "errors.json")  # Errors will be stored here

# Page range validation regex (supports formats like "1-3,5,8-10")
PAGE_RANGE_REGEX = r"^\d+(-\d+)?(,\d+(-\d+)?)*$"

def get_pdf_page_count(pdf_path):
    """Returns the total number of pages in a PDF."""
    try:
        with fitz.open(pdf_path) as doc:
            return len(doc)
    except Exception as e:
        print(f"❌ Error reading PDF '{pdf_path}': {e}")
        return None  # Return None if PDF cannot be read

def validate_page_range(page_range, total_pages):
    """Checks if the page range is valid within the given total pages."""
    invalid_pages = []
    for range_part in page_range.split(","):
        if "-" in range_part:
            start, end = map(int, range_part.split("-"))
            if start < 1 or end > total_pages or start > end:
                invalid_pages.append(range_part)
        else:
            page = int(range_part)
            if page < 1 or page > total_pages:
                invalid_pages.append(str(page))
    
    return invalid_pages

def validate_inputs():
    errors = {"missing_pdfs": [], "invalid_page_ranges": [], "empty_fields": [], "unreadable_pdfs": []}

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
                continue  # No need to check pages if PDF is missing

            # Validate that the PDF is readable and get total pages
            total_pages = get_pdf_page_count(pdf_path)
            if total_pages is None:
                errors["unreadable_pdfs"].append(f"Row {index+1}: Could not read {pdf_name}.")
                continue

            # Validate page range format
            if page_range and not re.match(PAGE_RANGE_REGEX, page_range):
                errors["invalid_page_ranges"].append(f"Row {index+1}: Invalid page range '{page_range}'.")
            else:
                # Ensure page range exists in the PDF
                invalid_pages = validate_page_range(page_range, total_pages)
                if invalid_pages:
                    errors["invalid_page_ranges"].append(
                        f"Row {index+1}: Pages {', '.join(invalid_pages)} out of range for {pdf_name} (Total: {total_pages} pages)."
                    )

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
