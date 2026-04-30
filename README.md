# 🦉 DHS Child Care Certificate Parser

A Streamlit app that parses Illinois DHS Child Care Certificate Report PDFs (form IL444-3492A) and exports a formatted Excel file.

## Features

- Upload a multi-page DHS Certificate Report PDF
- Automatically extracts provider info and all child records
- Displays an interactive data table with computed **Payment** column
  - Payment = Rate × Adj Days (if present) or Rate × Elig Days, minus Co-Pay
- One-click download of a formatted `.xlsx` file with two sheets:
  - **Children Detail** — all records + payment calculations
  - **Attendance Summary** — totals and attendance percentage

## Running Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/childcare-cert-parser.git
cd childcare-cert-parser

# 2. Create a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

## Deploying to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set main file to `app.py`
4. Click **Deploy**

No secrets or environment variables are required.

## File Structure

```
├── app.py             # Main Streamlit application
├── requirements.txt   # Python dependencies
└── README.md
```

## Notes

- The parser is tuned for the standard Illinois DHS IL444-3492A form layout.
- If your PDF has an unusual layout or scan quality, some fields may not parse automatically — the download will still work with whatever data was extracted.
- The app does not store or transmit any uploaded data.
