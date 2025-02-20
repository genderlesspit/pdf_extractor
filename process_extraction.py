import os
import csv
import fitz  # PyMuPDF

# Define paths
base_dir = os.path.expanduser("~/PDF_Extractor")
master_pdf_dir = os.path.join(base_dir, "Master_PDFs")
exports_dir = os.path.join(base_dir, "Exports")
grid_update_file = os.path.join(base_dir, "temp_grid_update.csv")

os.makedirs(exports_dir, exist_ok=True)

def extract_pages(pdf_name, page_ranges, output_name):
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

def process_extraction():
    if not os.path.exists(grid_update_file):
        print("❌ No input data found.")
        return

    with open(grid_update_file, "r", newline="", encoding="utf-8") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if not rows:
        print("❌ No valid rows found.")
        return

    for row in rows:
        try:
            pdf_name, output_name, page_ranges = row
            extract_pages(pdf_name, page_ranges, output_name)
        except ValueError:
            print("❌ Skipping invalid row.")

if __name__ == "__main__":
    process_extraction()
