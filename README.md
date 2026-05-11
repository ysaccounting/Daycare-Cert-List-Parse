# 🦉 DHS Child Care Certificate Parser

A web app that uses **Claude's vision AI** to read Illinois DHS Child Care Certificate Report PDFs (IL444-3492A) and export formatted Excel files.

## How it works

1. Upload a PDF (any number of pages)
2. Each page is converted to an image and sent to Claude claude-sonnet-4-20250514 via the Anthropic API
3. Claude reads the form fields directly — no regex, no OCR, no column guessing
4. Results display in an editable table; download as `.xlsx` with Payment calculations

## Deploy to Railway (recommended)

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub**
3. Add environment variable: `ANTHROPIC_API_KEY=sk-ant-...`
4. Railway auto-detects the Procfile and nixpacks.toml (installs poppler)
5. Deploy — done

## Run locally

```bash
# Install system dep (macOS)
brew install poppler

# Install system dep (Ubuntu)
sudo apt-get install poppler-utils

# Python setup
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Set API key
export ANTHROPIC_API_KEY=sk-ant-...

# Run
python app.py
# → http://localhost:5000
```

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | ✅ | Your Anthropic API key |
| `PORT` | auto | Set by Railway automatically |

## File structure

```
├── app.py              # Flask app + Claude parser + Excel builder
├── templates/
│   └── index.html      # Single-page UI
├── requirements.txt
├── Procfile            # gunicorn for Railway
├── nixpacks.toml       # installs poppler on Railway
└── .gitignore
```

## Notes

- Processing time: ~2–4 seconds per page (Claude vision API)
- A 27-page cert takes roughly 60–90 seconds
- The editable table lets you fix any errors before downloading
- Payment formula: `Rate × (Adj Days if present, else Elig Days) − Co-Pay`
