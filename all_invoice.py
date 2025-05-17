import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_invoice_data(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

        def extract(pattern, default=""):
            match = re.search(pattern, text)
            return match.group(1).strip() if match else default

        return {
            "payout_period": extract(r"Service Period\s*:\s*(\d{2}/\d{2}/\d{4} to \d{2}/\d{2}/\d{4})"),
            "file_name": os.path.basename(pdf_path),
            "fy_year": "2025-26",  # You may adjust logic here
            "year": "2025",
            "month": "04",
            "irn": extract(r"IRN\s*:\s*([a-f0-9]{64})"),
            "mann_gstin": extract(r"GSTIN\s*:\s*(29ABNFM9601R1Z9)"),
            "swiggy_gstin": extract(r"GSTIN\s*:\s*(29AAFCB7707D1ZQ)"),
            "sr_no": "1",
            "description": "Service Fee",
            "hsn": "996211",
            "unit_of_measure": "OTH",
            "quantity": "1",
            "unit_price": "2512.57",
            "base_amount": "2512.57",
            "discount": "0",
            "assessable_value": "2512.57",
            "cgst_rate": "9",
            "cgst_amount": "226.131",
            "sgst_rate": "9",
            "sgst_amount": "226.131",
            "igst_rate": "0",
            "igst_amount": "0",
            "comp_cess_rate": "0",
            "comp_cess_amount": "0",
            "state_cess_rate": "0",
            "state_cess_amount": "0",
            "total_amount": "2964.833",
            "other_charges_reimbursement_of_discount": extract(r"Other Charges.*?\n([\d,]+\.\d+)", "0"),
            "grand_total": extract(r"Grand Total\s+([\d,]+\.\d+)", "0").replace(",", ""),
            "brand_id": extract(r"Restaurant / Store ID\s*:\s*(\d+)"),
            "pan": extract(r"PAN\s*:\s*([A-Z0-9]+)"),
            "invoice_date": extract(r"Invoice Date\s*:\s*(\d{4}-\d{2}-\d{2})"),
            "invoice_number": extract(r"Invoice Number\s*:\s*([A-Z0-9]+)"),
            "original_invoice_number": extract(r"Original Invoice\s*No:\s*(.*?)\n"),
            "invoice_type": extract(r"Invoice Type\s*:\s*(\w+)")
        }
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return None

# Set folder path containing PDF files
folder_path = "C:/Users/91798/Downloads/Cleaning of Data & Merging into single excel/Cleaning of Data & Merging into single excel/Commission Invoices"
all_data = []

for file in os.listdir(folder_path):
    if file.lower().endswith(".pdf"):
        full_path = os.path.join(folder_path, file)
        data = extract_invoice_data(full_path)
        if data:
            all_data.append(data)

# Save to Excel
df = pd.DataFrame(all_data)
df.to_excel("all_invoices_output.xlsx", index=False)

print("✅ All invoice data extracted and saved to 'all_invoices_output.xlsx'")
