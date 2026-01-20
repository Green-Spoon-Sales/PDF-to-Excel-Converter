from pdfminer.high_level import extract_text
import pandas as pd
import re
import os
import glob

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = extract_text(pdf_path)
    return text

# Parsing function tailored to the provided text format
def parse_text_to_table(text):
    lines = text.split('\n')
    table_data = []
    payee = None
    invoice_number = None
    sold_to = ""
    sold_to_lines = []
    date = None

    for i, line in enumerate(lines):
        line = line.strip()
        if 'PAYEE:' in line and payee is None:
            payee = lines[i + 1].strip()
        elif 'INVOICE #' in line:
            invoice_number_match = re.search(r'INVOICE #(\d+)', line)
            if invoice_number_match:
                invoice_number = invoice_number_match.group(1).strip()
        elif 'SOLD TO:' in line:
            # Start capturing "SOLD TO" information
            sold_to_lines = [line.replace('SOLD TO:', '').strip()]
            # Continue to capture the next lines that belong to "SOLD TO"
            j = i + 1
            while j < len(lines) and lines[j].strip() and not re.match(r'^\d{12}', lines[j].strip()):
                sold_to_lines.append(lines[j].strip())
                j += 1
            sold_to = ' '.join(sold_to_lines).replace('TOL User : EMDROBOT', '').strip()
        elif re.match(r'^\d{1,2}/\d{2}/\d{2}$', line):  # Check for mm/dd/yy or m/dd/yy format
            date = line  # Assuming this line contains the date
        elif re.match(r'^\d{12}', line):  # Line starts with 12-digit UPC parts
            parts = re.split(r'\s+', line)
            if len(parts) < 10:
                continue  # Skip if the line does not have enough parts

            upc = parts[0]
            qty_ship = parts[1]
            # Determine where description ends by checking the remaining elements from the right
            description_end_idx = len(parts) - 7  # Description ends where the last 6 elements start
            description = ' '.join(parts[2:description_end_idx])
            nbr = parts[-6]
            date = parts[-5]
            comment = parts[-4]
            cost = parts[-3]
            disc = parts[-2]
            ext_cost = parts[-1]

            # Append the data for the current line to the table_data list
            table_data.append(
                [payee, sold_to, invoice_number, upc, qty_ship, description, nbr, date, comment, cost, disc, ext_cost])
        elif '-' in line:  # Handling lines with data separated by dashes
            parts = line.split('-')
            if len(parts) == 8:
                upc = parts[0].strip()
                qty_ship = parts[1].strip()
                description = parts[2].strip()
                nbr = parts[3].strip()
                date = parts[4].strip()
                comment = parts[5].strip()
                cost = parts[6].strip()
                disc = parts[7].strip()
                ext_cost = parts[8].strip()
                table_data.append(
                    [payee, sold_to, invoice_number, upc, qty_ship, description, nbr, date, comment, cost, disc, ext_cost])

    return table_data

# Main function to extract, parse, and write to Excel
def pdf_to_excel(pdf_path, output_excel_path):
    try:
        extracted_text = extract_text_from_pdf(pdf_path)
        parsed_data = parse_text_to_table(extracted_text)

        # Convert parsed data to a DataFrame
        df = pd.DataFrame(parsed_data,
                          columns=['Payee', 'Sold to', 'Invoice #', 'UPC#', 'QTY ship', 'Description', 'NBR', 'Date',
                                   'Comment', 'Cost', 'Disc $ or %', 'EXT-Cost'])

        # Write DataFrame to an Excel file
        df.to_excel(output_excel_path, index=False)
        print(f"Data has been written to {output_excel_path}")
    except Exception as e:
        print(f"Failed to process {pdf_path}: {e}")

# Function to process multiple PDFs in a directory
def process_multiple_pdfs(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + glob.glob(os.path.join(input_dir, '*.PDF'))
    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")
        # Define the output Excel file path
        output_excel_file = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_file))[0] + '_converted.xlsx')
        pdf_to_excel(pdf_file, output_excel_file)

# Example usage
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'
process_multiple_pdfs(input_dir, output_dir)
