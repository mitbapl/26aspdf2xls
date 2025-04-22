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

    for i, line in enumerate(lines):
        line = line.strip()

        # Match Deductor Summary (e.g., RULOANS... MUMR33583E 109367.00 5468.35 5468.35)
        if re.search(r"[A-Z]{4}\d{5}[A-Z]\s+\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}", line):
            match = re.search(r"(.+?)\s+([A-Z]{4}\d{5}[A-Z])\s+(\d+[.,]?\d*)\s+(\d+[.,]?\d*)\s+(\d+[.,]?\d*)", line)
            if match:
                current_deductor = match.group(1).strip()
                current_tan = match.group(2).strip()

        # Match Transaction Detail (e.g., 194H 23-Aug-2024 F 09-Nov-2024 - 53607.00 2680.35 2680.35)
        txn_match = re.match(
            r"(194[A-Z]?)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+[FMPUGZ]\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s*(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})",
            line
        )

        if txn_match and current_deductor and current_tan:
            try:
                entries.append({
                    "Deductor": current_deductor,
                    "TAN": current_tan,
                    "Section": txn_match.group(1),
                    "Transaction Date": txn_match.group(2),
                    "Amount Paid": float(txn_match.group(3).replace(',', '')),
                    "TDS Deducted": float(txn_match.group(4).replace(',', '')),
                    "TDS Deposited": float(txn_match.group(5).replace(',', '')),
                })
            except Exception as e:
                print(f"Skipping line {i} due to error: {e}")
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
