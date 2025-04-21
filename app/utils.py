import re
import os
from PyPDF2 import PdfReader
import pandas as pd

def extract_tds_entries_from_pdf(pdf_path):
    """
    Extracts TDS Part-I entries from 26AS PDF.
    """
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    # Pattern to match deductor level summary
    deductor_block = re.findall(
        r"Name of Deductor\s+(.*?)\s+TAN of Deductor\s+([A-Z]{4}\d{5}[A-Z])\s+Total Amount Paid/[\s\S]+?(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)",
        text
    )

    # Pattern to match transaction level detail
    transaction_block = re.findall(
        r"(\d{3}[A-Z]?)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+[FMUPZ]\s+\d{2}-[A-Za-z]{3}-\d{4}\s+-?\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)",
        text
    )

    records = []
    for i, (deductor, tan, amt, tds, deposited) in enumerate(deductor_block):
        # Map corresponding transaction(s)
        # Matching by order; ensure lists align
        if i < len(transaction_block):
            sec, date, tr_amt, tr_tds, tr_dep = transaction_block[i]
            records.append({
                "Deductor": deductor.strip(),
                "TAN": tan.strip(),
                "Section": sec.strip(),
                "Transaction Date": date.strip(),
                "Amount Paid": float(tr_amt),
                "TDS Deducted": float(tr_tds),
                "TDS Deposited": float(tr_dep)
            })
    return records

def convert_to_excel(records, output_excel_path):
    """
    Saves extracted TDS records to Excel.
    """
    df = pd.DataFrame(records)
    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
    df.to_excel(output_excel_path, index=False)
    print(f"Excel file saved: {output_excel_path}")

def process_pdf_to_excel(pdf_path, output_excel_path):
    records = extract_tds_entries_from_pdf(pdf_path)
    convert_to_excel(records, output_excel_path)
