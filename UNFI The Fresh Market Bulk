import os
import re
import glob
import pytesseract
import pandas as pd
from pdf2image import convert_from_path

# Set the Tesseract command path (adjust if necessary)
pytesseract.pytesseract.tesseract_cmd = '/opt/homebrew/bin/tesseract'

COLUMNS = [
    "Supplier Name",
    "Supplier Brand",
    "Customer Invoice",
    "Program Description",
    "Promo Period",
    "UPC Code",
    "Item Description",
    "Chargeback Quantity",
    "Chargeback Amount",
    "Total Chargeback",
    "Customer Notes"
]


def ocr_extract_text(file_path, dpi=300):
    """Convert PDF pages to images and perform OCR using Tesseract."""
    try:
        pages = convert_from_path(file_path, dpi=dpi, poppler_path='/opt/homebrew/bin')
        texts = []
        for page in pages:
            text = pytesseract.image_to_string(page, config="--psm 6")
            texts.append(text)
        return texts
    except Exception as e:
        print(f"Error converting PDF to images for {file_path}: {e}")
        return []


def parse_table_row(row_text):
    """Parse a single table row with debug output."""
    # Debug: Print the row being processed
    print(f"Processing row: {row_text}")

    # Remove extra whitespace and normalize spaces
    row_text = ' '.join(row_text.split())

    # Try to match the row pattern
    # Looking for: Supplier Name, Brand, Invoice, Program, Period, UPC, Description, Total, Notes
    pattern = r"""
        (RAO'S\s+SPECIALTY\s+FOODS[,\.]?\s*INC\.?)\s+  # Supplier Name
        (RAOS)\s+                                      # Brand
        (\d+)\s+                                       # Invoice Number
        (OTHER)\s+                                     # Program Description
        (\d{2}\.\d{2}\.\d{4})\s+                      # Promo Period
        (\d{11,12})\s+                                # UPC Code
        (.+?)\s+                                      # Item Description
        (\$\d+\.\d{2})\s*                            # Total Chargeback
        (.+)                                          # Customer Notes
    """

    match = re.match(pattern, row_text, re.VERBOSE)

    if match:
        print("Row matched pattern!")
        return [
            match.group(1).strip(),  # Supplier Name
            match.group(2).strip(),  # Brand
            match.group(3).strip(),  # Invoice
            match.group(4).strip(),  # Program
            match.group(5).strip(),  # Period
            match.group(6).strip(),  # UPC
            match.group(7).strip(),  # Description
            "",  # Chargeback Quantity (empty)
            "",  # Chargeback Amount (empty)
            match.group(8).strip(),  # Total
            match.group(9).strip()  # Notes
        ]

    print("Row did not match pattern")
    return None


def find_and_parse_tables(texts):
    """Find and parse tables from all pages of the PDF."""
    all_rows = []

    for page_num, text in enumerate(texts, 1):
        print(f"\nProcessing page {page_num}")
        print("=" * 50)
        print("Page content snippet:")
        print(text[:500])  # Print first 500 chars of the page

        lines = text.split('\n')
        table_start = -1

        # Find table header
        for i, line in enumerate(lines):
            if ('Supplier Name' in line and 'UPC Code' in line) or \
                    ('Supplier Name' in line and 'Customer Invoice' in line):
                table_start = i + 1
                print(f"Found table header at line {i}")
                break

        if table_start != -1:
            print("\nProcessing table rows:")
            # Process each line after the headers
            for line in lines[table_start:]:
                if line.strip():  # Only process non-empty lines
                    print(f"\nChecking line: {line}")
                    # More flexible row detection
                    if any(keyword in line for keyword in ["RAO", "SPECIALTY", "FOODS"]):
                        parsed_row = parse_table_row(line)
                        if parsed_row:
                            print("Successfully parsed row!")
                            all_rows.append(parsed_row)
                        else:
                            print("Failed to parse row")

    print(f"\nTotal rows parsed: {len(all_rows)}")
    return all_rows


def process_pdf_file(file_path, output_dir):
    """Process a single PDF file."""
    print(f"\nProcessing {file_path} with OCR...")
    texts = ocr_extract_text(file_path)
    if not texts:
        print(f"No text extracted from {file_path} after OCR.")
        return

    table_rows = find_and_parse_tables(texts)
    if not table_rows:
        print(f"No table rows detected in {file_path}. Check parsing logic.")
        return

    df = pd.DataFrame(table_rows, columns=COLUMNS)
    base = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(output_dir, f"{base}_converted.xlsx")

    try:
        df.to_excel(output_file, index=False)
        print(f"Excel file created: {output_file}")
    except Exception as e:
        print(f"Error writing {output_file}: {e}")


def process_multiple_pdfs(input_dir, output_dir):
    """Process all PDFs in the input directory."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + \
                glob.glob(os.path.join(input_dir, '*.PDF'))

    print(f"Found {len(pdf_files)} PDF files in {input_dir}")
    for file_path in pdf_files:
        process_pdf_file(file_path, output_dir)


# Directory paths
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'

# Process all PDFs
process_multiple_pdfs(input_dir, output_dir)
