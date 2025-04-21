import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_tds_entries_from_pdf(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    lines = text.splitlines()

    entries = []
    current_deductor = ""
    current_tan = ""

    for i in range(len(lines)):
        # Match Deductor and TAN block
        if re.search(r"TAN of Deductor", lines[i]):
            # Find the line before that contains deductor name
            current_deductor = lines[i - 1].strip()
            tan_match = re.search(r'([A-Z]{4}\d{5}[A-Z])', lines[i])
            current_tan = tan_match.group(1) if tan_match else ""

        # Match transaction line
        txn_match = re.match(
            r'^(1[9][0-9][A-Z]?)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+[FMUPZ]\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s+(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)',
            lines[i]
        )
        if txn_match:
            section = txn_match.group(1)
            date = txn_match.group(2)
            amount = float(txn_match.group(3).replace(',', ''))
            tds = float(txn_match.group(4).replace(',', ''))
            deposited = float(txn_match.group(5).replace(',', ''))

            entries.append({
                "Deductor": current_deductor,
                "TAN": current_tan,
                "Section": section,
                "Transaction Date": date,
                "Amount Paid": amount,
                "TDS Deducted": tds,
                "TDS Deposited": deposited
            })

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
