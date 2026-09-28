# GSTR-9 Auto Fill and AIS JSON Extractor

Automated GSTR-9 format filler. Upload a zip with Speqta reports + monthly PDFs → Download filled Excel.

It also includes an **AIS JSON Extractor** for readable JSON downloads from the
Income Tax AIS portal. Upload the JSON and download an Excel workbook with an
index plus a worksheet for every reported-data category found in the file.

> If the portal file is encrypted or cannot be opened as JSON, open/export it
> with the official AIS Utility first, then upload the readable JSON export.

## Deploy to Streamlit Cloud

1. Upload this folder to a GitHub repository
2. Go to https://share.streamlit.io
3. Click "New app" → select your repo → set main file = `app.py`
4. Click Deploy

## Extract AIS reported data

1. Open the app and choose **AIS JSON Extractor**.
2. Upload the readable `.json` file downloaded/exported from the AIS portal.
3. Select **Extract reported data** and download `AIS_REPORTED_DATA.xlsx`.

The extractor does not depend on fixed AIS category names; each list of JSON
objects is retained as a separate worksheet, including newly added categories.

Your team gets a permanent URL like:
`https://your-app-name.streamlit.app`

## What goes in the zip

- Full Year 3B Report.xlsx
- GSTR-3B Sales Summary.xlsx
- GSTR-3B ITC Summary.xlsx
- GSTR-1 Sales Summary.xlsx
- All GSTR-3B monthly PDFs (filenames must contain "GSTR3B")
- All GSTR-1 monthly PDFs (filenames must contain "GSTR1")
