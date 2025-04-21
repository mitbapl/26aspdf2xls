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

        if not text:
            print("No text extracted from PDF.")
            raise ValueError("Failed to extract any text from the PDF.")

        # Regex pattern to extract relevant data
        tds_pattern = re.compile(r"(?P<deductor>.+?)\s+(?P<tan>MUM[A-Z0-9]+)\s+(?P<amount_paid>\d+\.\d+)\s+(?P<tds_deducted>\d+\.\d+)")
        matches = tds_pattern.findall(text)

        # Parse matched data into a list of dictionaries
        data = []
        for match in matches:
            data.append({
                "Deductor": match[0],
                "TAN": match[1],
                "Amount Paid": float(match[2]),
                "TDS Deducted": float(match[3]),
            })
        
        if not data:
            print("No TDS data found in PDF.")
            raise ValueError("No TDS data found in the PDF.")

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
        df["Sub Total"] = df["TDS Deducted"].sum()

        print(f"Extracted data: {df.head()}")  # Log a sample of extracted data

        # Write data to Excel
        output_dir = os.path.dirname(output_excel)
        if not os.path.exists(output_dir):
            print(f"Creating output directory: {output_dir}")
            os.makedirs(output_dir)

        # Use Excel writer to save data
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="TDS Details")

            # Add a total row
            worksheet = writer.sheets["TDS Details"]
            worksheet.write(len(df) + 1, 0, "Total")
            worksheet.write(len(df) + 1, 3, df["TDS Deducted"].sum())

        print(f"Excel file generated at: {output_excel}")

    except Exception as e:
        print(f"Error during PDF to Excel conversion: {str(e)}")
        raise
