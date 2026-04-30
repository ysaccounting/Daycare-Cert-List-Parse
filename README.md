# 🦉 DHS Child Care Certificate Parser

A Streamlit app that parses Illinois DHS Child Care Certificate Report PDFs (form IL444-3492A) using OCR and exports a formatted Excel file.

## Features

- Upload a multi-page DHS Certificate Report PDF (scanned or digital)
- OCR extracts provider info and all child records using positional analysis
- **Editable table** — fix any OCR errors before downloading
- Computes **Payment** column: Rate × Adj Days (if present) or Rate × Elig Days, minus Co-Pay
- One-click Excel download named after the daycare and period (e.g. `Wise Owl Oct 2025 Certs.xlsx`)
- Two-sheet Excel: **Children Detail** + **Attendance Summary**

## System Dependency

This app requires **Tesseract OCR** to be installed on the host system.

### Streamlit Cloud
Add a `packages.txt` file to the repo root with:
```
tesseract-ocr
poppler-utils
```

### Local (Ubuntu/Debian)
```bash
sudo apt-get install tesseract-ocr poppler-utils
```

### Local (macOS)
```bash
brew install tesseract poppler
```

### Local (Windows)
Install [Tesseract for Windows](https://github.com/UB-Mannheim/tesseract/wiki) and [poppler for Windows](https://github.com/oschwartz10612/poppler-windows).

## Running Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/childcare-cert-parser.git
cd childcare-cert-parser

# 2. Install system dependencies (see above)

# 3. Create a virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 4. Install Python dependencies
pip install -r requirements.txt

# 5. Run the app
streamlit run app.py
```

## Deploying to Streamlit Cloud

1. Push this repo to GitHub (include `packages.txt`)
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select repo → set main file to `app.py`
4. Click **Deploy** — Streamlit Cloud will install Tesseract automatically via `packages.txt`

## File Structure

```
├── app.py            # Main Streamlit application
├── requirements.txt  # Python dependencies
├── packages.txt      # System packages (Tesseract, poppler)
└── README.md
```

## Notes

- OCR quality depends on scan quality. The editable table lets you correct any errors before export.
- The column position detector is calibrated for the standard IL444-3492A layout at 300 DPI.
- No data is stored or transmitted — everything runs locally in your browser session.
