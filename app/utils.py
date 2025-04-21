import re
import pandas as pd
from PyPDF2 import PdfReader
import os

def extract_tds_data_from_pdf(pdf_path):
    """
    Extract TDS data from a PDF file.
    """
    print(f"Extracting TDS data from PDF: {pdf_path}")

    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""

        if not text.strip():
            raise ValueError("No text extracted from the PDF.")

        # Print extracted text for debugging
        print("Extracted Text (first 500 characters):")
        print(text[:500])

        # Regex pattern to extract TDS details
        tds_pattern = re.compile(
            r"(?P<deductor>[A-Z\s]+)\s+(?P<tan>MUM[A-Z0-9]+)\s+(?P<amount_paid>\d+\.\d+)\s+(?P<tds_deducted>\d+\.\d+)"
        )
        matches = tds_pattern.findall(text)

        if not matches:
            raise ValueError("No TDS data found in the PDF. Please check the data format.")

        # Parse matched data into a list of dictionaries
        data = [
            {
                "Deductor": match[0].strip(),
                "TAN": match[1].strip(),
                "Amount Paid": float(match[2]),
                "TDS Deducted": float(match[3]),
            }
            for match in matches
        ]

        return data

    except Exception as e:
        print(f"Error extracting TDS data: {str(e)}")
        raise

def process_pdf_to_excel(pdf_path, output_excel):
    """
    Process a PDF file and export TDS data to an Excel file.
    """
    print(f"Generating Excel file at: {output_excel}")
    print(f"Processing {pdf_path} and saving to {output_excel}")

    try:
        # Extract TDS data
        data = extract_tds_data_from_pdf(pdf_path)
        if not data:
            print("No data extracted. Excel file will not be generated.")
            return

        # Convert data to a DataFrame
        df = pd.DataFrame(data)
        print("Extracted DataFrame:")
        print(df)

        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_excel)
        os.makedirs(output_dir, exist_ok=True)

        # Write data to Excel
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="TDS Details")

            # Add a total row
            worksheet = writer.sheets["TDS Details"]
            worksheet.write(len(df), 0, "Total")
            worksheet.write(len(df), 3, df["TDS Deducted"].sum())

        print(f"Excel file successfully generated at: {output_excel}")

    except Exception as e:
        print(f"Error during PDF to Excel conversion: {str(e)}")
        raise
