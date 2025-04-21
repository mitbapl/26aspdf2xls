import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_tds_entries_from_pdf(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = "\n".join([page.extract_text() or "" for page in reader.pages])
    lines = text.splitlines()

    entries = []
    current_deductor = None
    current_tan = None

    for i in range(len(lines)):
        line = lines[i].strip()

        # Detect deductor block
        if line.isupper() and "LIMITED" in line or "PRIVATE" in line or "BANK" in line:
            if i + 1 < len(lines) and "TAN of Deductor" in lines[i + 1]:
                current_deductor = line.strip()
                tan_match = re.search(r'([A-Z]{4}\d{5}[A-Z])', lines[i + 1])
                current_tan = tan_match.group(1) if tan_match else None

        # Detect transaction line
        txn_match = re.match(
            r"^(1[9][0-9][A-Z]?)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+[FMUPZ]\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s*(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)",
            line
        )
        if txn_match and current_deductor and current_tan:
            try:
                section = txn_match.group(1)
                txn_date = txn_match.group(2)
                amount = float(txn_match.group(3).replace(',', ''))
                tds_deducted = float(txn_match.group(4).replace(',', ''))
                tds_deposited = float(txn_match.group(5).replace(',', ''))

                entries.append({
                    "Deductor": current_deductor,
                    "TAN": current_tan,
                    "Section": section,
                    "Transaction Date": txn_date,
                    "Amount Paid": amount,
                    "TDS Deducted": tds_deducted,
                    "TDS Deposited": tds_deposited
                })
            except ValueError:
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
