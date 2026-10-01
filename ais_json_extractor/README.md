# AIS JSON Extractor

A standalone Streamlit tool that converts a readable JSON download from the Income Tax AIS portal into an Excel workbook of reported data.

## Use

1. Run `streamlit run app.py` from this directory.
2. Upload the readable AIS `.json` file.
3. Select **Extract reported data** and download `AIS_REPORTED_DATA.xlsx`.

The workbook has an **Index** worksheet and a separate worksheet for every JSON list of reported records found in the file. Nested fields are preserved as dot-separated columns.

> If an AIS download is encrypted or is not readable JSON, open or export it with the official AIS Utility first. This tool processes readable JSON only.

## Install

```bash
pip install -r requirements.txt
streamlit run app.py
```
