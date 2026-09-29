# GSTR-9 and AIS Tools

This Streamlit application provides two workflows from the same deployed entry point:

- **GSTR-9 Auto Fill** — upload a ZIP containing Speqta reports and monthly PDFs, then download a filled GSTR-9 workbook.
- **AIS JSON Extractor** — upload a readable AIS JSON export from the Income Tax portal, then download an Excel workbook with an index and one worksheet for each reported-data category.

## Deploy to Streamlit Cloud

1. Upload the `gstr9_streamlit` folder to a GitHub repository.
2. In Streamlit Cloud, create an app with main file `gstr9_streamlit/app.py` (or `app.py` when this folder is the repository root).
3. Deploy. Use the **Choose a tool** selector in the deployed app to switch between GSTR-9 and AIS workflows.

## AIS JSON Extractor

Upload the readable `.json` file downloaded from the AIS portal and select **Extract reported data**. The workbook preserves nested fields as dot-separated columns and includes every discovered list of reported JSON records.

> Encrypted or non-JSON AIS downloads must first be opened or exported with the official AIS Utility. This application accepts readable UTF-8 JSON only.

## GSTR-9 ZIP contents

- Full Year 3B Report.xlsx
- GSTR-3B Sales Summary.xlsx
- GSTR-3B ITC Summary.xlsx
- GSTR-1 Sales Summary.xlsx
- All GSTR-3B monthly PDFs (filenames must contain "GSTR3B")
- All GSTR-1 monthly PDFs (filenames must contain "GSTR1")
