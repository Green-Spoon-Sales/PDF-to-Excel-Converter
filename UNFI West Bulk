import pdfplumber
import pandas as pd
import re
import os
import glob


# Function to extract data from a PDF file
def extract_data_from_pdf(pdf_path, brands_to_extract):
    data = []
    week_ending = None  # Initialize week_ending variable
    customer_info = None  # Initialize customer_info variable to persist across pages

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')

            # Extract the week ending date
            if not week_ending:
                for line in lines:
                    if "Week ending" in line:
                        week_ending = line.split(':')[-1].strip()
                        break

            # Extract customer information and the table data
            for line in lines:
                if line.startswith("Customer :"):
                    customer_info = line.split("Customer :")[1].strip()

                if any(brand in line.upper() for brand in brands_to_extract) and week_ending and customer_info:
                    # Debug print to check the line content
                    print(f"Processing line: {line}")

                    # Split the line by whitespace
                    columns = line.split()

                    # Find the indices for known columns to handle description extraction
                    try:
                        invoice_idx = next(i for i, col in enumerate(columns) if col.isdigit() and len(col) == 9)
                        ordered_idx = invoice_idx + 1
                        shipped_idx = ordered_idx + 1
                        whlse_idx = shipped_idx + 1
                        total_discount_idx = whlse_idx + 1
                        mcb_percentage_idx = total_discount_idx + 1
                        mcb_idx = mcb_percentage_idx + 1

                        brand = columns[0]
                        product = columns[1]
                        unit = columns[2]

                        # Capture the unit with OZ if it's part of the unit
                        if not unit.endswith("OZ") and len(columns) > 3:
                            unit += " " + columns[3]
                            description_start_index = 4
                        else:
                            description_start_index = 3

                        description = " ".join(columns[description_start_index:invoice_idx])
                        invoice = columns[invoice_idx]
                        ordered = columns[ordered_idx]
                        shipped = columns[shipped_idx]
                        whlse = columns[whlse_idx]
                        total_discount = columns[total_discount_idx]
                        mcb_percentage = columns[mcb_percentage_idx]
                        mcb = columns[mcb_idx]

                        data.append([
                            week_ending, customer_info, brand, product, unit,
                            description, invoice, ordered, shipped, whlse,
                            total_discount, mcb_percentage, mcb
                        ])

                        # Debug print to check extracted data
                        print(f"Extracted data: {data[-1]}")
                    except StopIteration:
                        print(f"Failed to parse line: {line}")

    return data


# Function to convert PDF to Excel
def pdf_to_excel(pdf_path, output_excel_path, brands_to_extract):
    data = extract_data_from_pdf(pdf_path, brands_to_extract)

    if not data:
        print(f"No data extracted from {pdf_path}. Please check the PDF content and extraction logic.")
        return

    # Create a DataFrame
    columns = ["Week Ending", "Customer Info", "Brand", "Product", "Unit", "Description", "Invoice",
               "Ordered", "Shipped", "Whlse", "Total Discount %", "MCB %", "MCB"]

    df = pd.DataFrame(data, columns=columns)

    # Save to Excel
    df.to_excel(output_excel_path, index=False)
    print(f"Data has been written to {output_excel_path}")


# Function to process multiple PDFs in a directory
def process_multiple_pdfs(input_dir, output_dir, brands_to_extract):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + glob.glob(os.path.join(input_dir, '*.PDF'))
    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")
        # Define the output Excel file path
        output_excel_file = os.path.join(output_dir,
                                         os.path.splitext(os.path.basename(pdf_file))[0] + '_converted.xlsx')
        pdf_to_excel(pdf_file, output_excel_file, brands_to_extract)


# Example usage
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'
brands_to_extract = ["RAOS", "NOOSA", "*NOOSA", "MICHAEL ANGELO'S", "PACIFIC", "LATE JULY", "PACIFIC", "@RAOS"]
process_multiple_pdfs(input_dir, output_dir, brands_to_extract)
