import os
import re
import csv
import pdfplumber
from datetime import datetime
from openpyxl import Workbook

# Folder with PDF files
pdf_folder = 'AMEX'

# Function to extract fields from text using pdfplumber
def extract_patient_data(text, filename):
    data = {
        "Filename": filename,
        "Patient Name": "",
        "Sex": "",
        "Age": "",
        "Hospital No.": "",
        "Date": ""
    }

    # Match name, sex, age on same or multiline
    match = re.search(
        r"Patient(?:’s)? Name:\s*([A-Za-z ,.'\-]+)\s*Sex:\s*(Female|Male|F|M)\s*Age:\s*(\d{1,3})",
        text, re.IGNORECASE | re.DOTALL
    )
    if match:
        data["Patient Name"] = match.group(1).strip()
        data["Sex"] = match.group(2).capitalize()
        data["Age"] = match.group(3).strip()

    # Match Hospital No.
    hosp_match = re.search(r"Hospital No\.\s*:\s*([A-Za-z0-9\-]+)", text)
    if hosp_match:
        data["Hospital No."] = hosp_match.group(1).strip()

    # Match Date
    date_match = re.search(r"Date:\s*(\d{2}/\d{2}/\d{4})", text)
    if date_match:
        data["Date"] = date_match.group(1).strip()

    return data

# Extract text using pdfplumber
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Process all PDFs
records = []
for file in os.listdir(pdf_folder):
    if file.lower().endswith(".pdf"):
        path = os.path.join(pdf_folder, file)
        print(f"Processing: {file}")
        text = extract_text_from_pdf(path)
        record = extract_patient_data(text, file)
        records.append(record)

# Save output
fields = ["Filename", "Patient Name", "Sex", "Age", "Hospital No.", "Date"]
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
excel_filename = f"all_patient_data_{timestamp}.xlsx"
# csv_filename = f"all_patient_data_{timestamp}.csv"

# Excel
wb = Workbook()
ws = wb.active
ws.title = "Patient Records"
ws.append(fields)
for record in records:
    ws.append([record.get(f, "") for f in fields])
wb.save(excel_filename)

# CSV
# with open(csv_filename, "w", newline="") as f:
#     writer = csv.DictWriter(f, fieldnames=fields)
#     writer.writeheader()
#     for record in records:
#         writer.writerow({key: record.get(key, "") for key in fields})

print("\n✅ Extraction complete!")
print(f"→ Excel saved as: {excel_filename}")
# print(f"→ CSV saved as:   {csv_filename}")
