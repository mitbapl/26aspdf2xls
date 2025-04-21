import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_tds_entries_from_pdf(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = "\n".join([page.extract_text() or "" for page in reader.pages])

    # 1. Extract Deductor Blocks
    deductor_pattern = re.compile(
        r"(?P<deductor>[A-Z0-9\s\.\-&()]+?)\s+(?P<tan>[A-Z]{4}\d{5}[A-Z])\s+(?P<amount_paid>-?\d[\d,]*\.\d{2})\s+(?P<tds_deducted>-?\d[\d,]*\.\d{2})\s+(?P<tds_deposited>-?\d[\d,]*\.\d{2})"
    )
    deductors = deductor_pattern.findall(text)

    # 2. Extract Transaction Rows (need to assign to the most recent deductor)
    txn_pattern = re.compile(
        r"(?P<section>194[A-Z]*)\s+(?P<date>\d{2}-[A-Za-z]{3}-\d{4})\s+[FMPUGZ]\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s*(?P<amount>-?\d[\d,]*\.\d{2})\s+(?P<tds>-?\d[\d,]*\.\d{2})\s+(?P<deposited>-?\d[\d,]*\.\d{2})"
    )
    transactions = txn_pattern.findall(text)

    # 3. Stitch Together
    results = []
    current_deductor = None
    current_tan = None

    # Walk through lines to align deductors + transactions properly
    lines = text.splitlines()
    for i in range(len(lines)):
        # Match deductor block in-line
        match = deductor_pattern.match(lines[i])
        if match:
            current_deductor = match.group('deductor').strip()
            current_tan = match.group('tan').strip()
            continue

        # Match transaction lines
        txn_match = txn_pattern.match(lines[i])
        if txn_match and current_deductor and current_tan:
            try:
                results.append({
                    "Deductor": current_deductor,
                    "TAN": current_tan,
                    "Section": txn_match.group('section'),
                    "Transaction Date": txn_match.group('date'),
                    "Amount Paid": float(txn_match.group('amount').replace(',', '')),
                    "TDS Deducted": float(txn_match.group('tds').replace(',', '')),
                    "TDS Deposited": float(txn_match.group('deposited').replace(',', '')),
                })
            except Exception as e:
                print(f"Skipping row due to error: {e}")
                continue

    return results

def convert_to_excel(records, output_excel_path):
    df = pd.DataFrame(records)
    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
    df.to_excel(output_excel_path, index=False)

def process_pdf_to_excel(pdf_path, output_excel_path):
    records = extract_tds_entries_from_pdf(pdf_path)
    if not records:
        raise ValueError("No TDS data extracted from PDF.")
    convert_to_excel(records, output_excel_path)
