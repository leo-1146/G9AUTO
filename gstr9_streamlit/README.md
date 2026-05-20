# GSTR-9 Auto Fill

Automated GSTR-9 format filler. Upload a zip with Speqta reports + monthly PDFs → Download filled Excel.

## Deploy to Streamlit Cloud

1. Upload this folder to a GitHub repository
2. Go to https://share.streamlit.io
3. Click "New app" → select your repo → set main file = `app.py`
4. Click Deploy

Your team gets a permanent URL like:
`https://your-app-name.streamlit.app`

## What goes in the zip

- Full Year 3B Report.xlsx
- GSTR-3B Sales Summary.xlsx
- GSTR-3B ITC Summary.xlsx
- GSTR-1 Sales Summary.xlsx
- All GSTR-3B monthly PDFs (filenames must contain "GSTR3B")
- All GSTR-1 monthly PDFs (filenames must contain "GSTR1")
