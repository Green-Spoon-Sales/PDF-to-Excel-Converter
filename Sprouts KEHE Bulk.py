import re
import PyPDF2
import pandas as pd
import os

# Define the pattern for parsing the lines
pattern = re.compile(
    r"(?P<code>\w+)\s+(?P<upc>\d+)\s+(?P<description>.*?)\s+\$(?P<unit_price>[\d\.]+)\s+(?P<start_date>\d{2}/\d{2}/\d{4})\s+(?P<end_date>\d{2}/\d{2}/\d{4})\s+(?P<units>\d+)\s+(?P<quantity>\d+)\s+\$(?P<total_price>[\d\.]+)"
)


def process_line(line):
    match = pattern.match(line)
    if match:
        return match.groupdict()
    return None


def parse_pdf(file_path):
    # Extract text from PDF
    parsed_data = []
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            page_text = page.extract_text()
            lines = page_text.splitlines()

            # Process each line, including the first line of each page
            for line in lines:
                parsed_line = process_line(line)
                if parsed_line:
                    print(f"All parts matched successfully!\n{parsed_line}")
                    parsed_data.append(parsed_line)
                else:
                    print(f"No match for line: {line}")

    return parsed_data


def export_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Data exported successfully to {output_file}")


def process_pdfs_in_folder(input_folder, output_folder):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Process each PDF in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith('.pdf'):
            pdf_file_path = os.path.join(input_folder, file_name)
            parsed_data = parse_pdf(pdf_file_path)

            if parsed_data:
                # Create output Excel file path
                excel_file_name = file_name.replace('.pdf', '.xlsx')
                excel_file_path = os.path.join(output_folder, excel_file_name)
                export_to_excel(parsed_data, excel_file_path)
            else:
                print(f"No valid data was parsed from {file_name}.")


if __name__ == "__main__":
    # Set the input and output folder paths
    input_folder = "/Users/joewilt/Desktop/Incoming Sovos files"  # Input folder path
    output_folder = "/Users/joewilt/Desktop/Converted Sovos files"  # Output folder path

    # Process all PDFs in the input folder and export to the output folder
    process_pdfs_in_folder(input_folder, output_folder)
