import os
import csv
import fitz  # PyMuPDF for PDF manipulation
import json
import time

# Define paths
base_dir = os.path.expanduser("~/PDF_Extractor")
master_pdf_dir = os.path.join(base_dir, "Master_PDFs")
exports_dir = os.path.join(base_dir, "Exports")
grid_update_file = os.path.join(base_dir, "temp_grid_update.csv")
failure_log_file = os.path.join(base_dir, "failed_extractions.json")

os.makedirs(exports_dir, exist_ok=True)

MAX_RETRIES = 3  # Number of times to retry failed extractions

def extract_pages(pdf_name, page_ranges, output_name):
    """Extracts non-contiguous pages from a PDF and saves them as a new PDF."""
    pdf_path = os.path.join(master_pdf_dir, pdf_name)
    output_path = os.path.join(exports_dir, f"{output_name}.pdf")

    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_name}")
        return False

    try:
        doc = fitz.open(pdf_path)
        new_doc = fitz.open()

        # Process page ranges (e.g., "1-3,5,8-10")
        for range_part in page_ranges.split(","):
            if "-" in range_part:
                start, end = map(int, range_part.split("-"))
                for page_num in range(start, end + 1):
                    if 1 <= page_num <= len(doc):
                        new_doc.insert_pdf(doc, from_page=page_num-1, to_page=page_num-1)
            else:
                page_num = int(range_part)
                if 1 <= page_num <= len(doc):
                    new_doc.insert_pdf(doc, from_page=page_num-1, to_page=page_num-1)

        # Save the extracted pages to a new PDF
        new_doc.save(output_path)
        new_doc.close()
        doc.close()
        print(f"✅ Extracted pages {page_ranges} from {pdf_name} → {output_name}.pdf")
        return True

    except Exception as e:
        print(f"❌ Error extracting pages from {pdf_name}: {e}")
        return False

def retry_failed_extractions():
    """Retries failed extractions up to MAX_RETRIES times."""
    if not os.path.exists(failure_log_file):
        return

    try:
        with open(failure_log_file, "r", encoding="utf-8") as file:
            failed_extractions = json.load(file)
    except json.JSONDecodeError:
        print("⚠ Error reading failed_extractions.json. Skipping retry process.")
        return

    if not failed_extractions:
        return

    print(f"🔄 Retrying {len(failed_extractions)} failed extractions...")

    new_failed_extractions = {}

    for key, attempt in failed_extractions.items():
        pdf_name, output_name, page_ranges = key.split("|")

        print(f"🔁 Attempting retry {attempt + 1}/{MAX_RETRIES} for: {output_name}")
        success = extract_pages(pdf_name, page_ranges, output_name)

        if not success:
            if attempt + 1 < MAX_RETRIES:
                new_failed_extractions[key] = attempt + 1  # Increment retry count
            else:
                print(f"❌ Maximum retries reached for: {output_name}. Skipping permanently.")

    # Update the failed extractions log
    with open(failure_log_file, "w", encoding="utf-8") as file:
        json.dump(new_failed_extractions, file, indent=4)

    if not new_failed_extractions:
        os.remove(failure_log_file)  # Clean up if no more failures

def process_extraction():
    """Processes the extraction workflow and handles failures."""
    if not os.path.exists(grid_update_file):
        print("❌ No input data found.")
        return

    with open(grid_update_file, "r", newline="", encoding="utf-8") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if not rows:
        print("❌ No valid rows found.")
        return

    failed_extractions = {}

    for row in rows:
        try:
            pdf_name, output_name, page_ranges = row
            success = extract_pages(pdf_name, page_ranges, output_name)

            if not success:
                failed_extractions[f"{pdf_name}|{output_name}|{page_ranges}"] = 0  # Start at 0 attempts

        except ValueError:
            print("❌ Skipping invalid row.")

    # Log failures to JSON for retry
    with open(failure_log_file, "w", encoding="utf-8") as file:
        json.dump(failed_extractions, file, indent=4)

    # Retry failed extractions
    retry_failed_extractions()

if __name__ == "__main__":
    process_extraction()
