import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_tds_data_from_pdf(pdf_path):
    """
    Extract TDS data from a PDF file.
    """
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

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
    return data

def process_pdf_to_excel(pdf_path, output_excel):
    print(f"Generating Excel file at: {output_excel}")
    # Rest of the code...

    """
    Process a PDF file and export TDS data to an Excel file.
    """
    data = extract_tds_data_from_pdf(pdf_path)
    if not data:
        return

    # Convert data to a DataFrame
    df = pd.DataFrame(data)
    df["Sub Total"] = df["TDS Deducted"].sum()

    # Write data to Excel
    writer = pd.ExcelWriter(output_excel, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="TDS Details")

    # Add a total row
    worksheet = writer.sheets["TDS Details"]
    worksheet.write(len(df) + 1, 0, "Total")
    worksheet.write(len(df) + 1, 3, df["TDS Deducted"].sum())

    writer.save()
