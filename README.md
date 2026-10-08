# Implant BOM Generator

A small Streamlit app for creating a concise implant BOM. Product catalogs are stored in editable, brand-specific Excel workbooks under `catalogs/`.

## Run locally on Windows

1. Install Python 3.11 or newer.
2. Open PowerShell in this folder.
3. Install the app packages:

   ```powershell
   py -m pip install -r requirements.txt
   ```

4. Start the app:

   ```powershell
   py -m streamlit run app.py --server.address 127.0.0.1
   ```

5. Open the local URL Streamlit prints, usually `http://localhost:8501`.

## Catalog workbooks

- `catalogs/ZimVie_Dental.xlsx`
- `catalogs/Dentsply_Sirona.xlsx`

Each workbook has separate sheets for implants, drills, drivers, and applicable prosthetic components. Keep column names and serialized list fields (such as `standard` and `denseExtra`) intact when editing. The app reads both workbooks when it starts; restart the app after editing them. Each brand selected in a case has its own short BOM Excel download.

The workbooks were seeded from the original `BOM 6OCT2026.txt` source. Protocols, placeholder part numbers, and pricing need review against current controlled source data before operational use. This app does not persist case data; the generated Excel files are the saved outputs.

## Deployment

The app can run on a private server or an approved cloud host that supports Python and Streamlit. Users visit the host's URL in a browser. Keep the app and catalog workbooks together, and configure authentication/network access before using real case information. A hosted app is not automatically private just because it is written in Python.