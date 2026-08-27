# NSE Disclosure Extractor

A Python tool that downloads recent NSE disclosures, processes PDFs, identifies financially relevant announcements, and creates Excel and JSON reports.

# Setup (powershell)

pip install playwright pandas openpyxl PyMuPDF pypdf
playwright install chromium

# Run 

Add company names to column B of `companies.xlsx`, then run:

python NSE.py


Results are saved in `outputs`.
