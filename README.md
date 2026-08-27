# NSE Disclosure Extractor

A Python automation tool that collects recent company disclosures from the National Stock Exchange of India, downloads the associated PDF documents, identifies financially relevant announcements, and generates structured Excel reports.

## Features

- Reads the company list from `companies.xlsx`
- Collects company-wise NSE disclosures for a configurable lookback period
- Downloads and processes disclosure PDFs
- Classifies financially relevant announcements and extracts supporting details
- Avoids duplicate documents and creates separate audited reports for each company

## Setup

Install Python and Google Chrome, then run:

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Run

Add or update company names in `companies.xlsx`, then run:

```powershell
python NSE.py
```

Results are created inside a timestamped folder under `outputs`. Press `Ctrl+C` to stop the current process.

## Notes

- Keep the browser open while collection is running.
- The saved browser profile and generated outputs are local-only and excluded from Git.
- Review generated classifications before relying on them for financial decisions.
