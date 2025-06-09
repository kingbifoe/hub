# 📘 PDF to Excel - Patient Report Extractor

This tool extracts structured patient data from radiology PDF reports and exports them to Excel and CSV.

---

## 🔍 Extracted Fields
- Patient Name
- Sex
- Age
- Hospital Number
- Date
- Filename

---

## 🗂 Folder Setup
Place all your `.pdf` files into the `/pdfs` folder.

---

## 💾 Output
The following files will be generated with timestamps:
- `all_patient_data_YYYYMMDD_HHMMSS.xlsx`
- `all_patient_data_YYYYMMDD_HHMMSS.csv`

---

## ✅ How to Use

```bash
python pdf_to_excel.py
