import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_tds_entries_from_pdf(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = "\n".join([page.extract_text() or "" for page in reader.pages])

    # Extract all matching transaction entries
    matches = re.findall(
        r"(194[A-Z])\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+F\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
        text
    )

    entries = []
    for match in matches:
        try:
            entries.append({
                "Section": match[0],
                "Transaction Date": match[1],
                "Amount Paid": float(match[2].replace(',', '')),
                "TDS Deducted": float(match[3].replace(',', '')),
                "TDS Deposited": float(match[4].replace(',', '')),
            })
        except Exception as e:
            print(f"Skipping row due to error: {e}")
            continue

    return entries

def convert_to_excel(records, output_excel_path):
    df = pd.DataFrame(records)
    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
    df.to_excel(output_excel_path, index=False)

def process_pdf_to_excel(pdf_path, output_excel_path):
    records = extract_tds_entries_from_pdf(pdf_path)
    if not records:
        raise ValueError("No TDS data extracted from PDF.")
    convert_to_excel(records, output_excel_path)
