# SBI Card PDF Suite

Two-in-one web app:
1. **Unlock PDFs** — batch password removal (multiple passwords supported via retry)
2. **Consolidate to Excel** — extract all transactions from SBI Card statements into one sorted Excel file

## Features

- Upload 1–50 PDFs at once
- Unlock with one password; retry different passwords for remaining locked files
- Automatically detects text-based vs scanned PDFs (OCR fallback via Tesseract)
- Extracts: Date, Description, Debit, Credit, Source File
- Excel output: All Transactions sheet (sorted by date) + Summary + Source Details

## System Requirements

```bash
# Ubuntu/Debian
sudo apt install tesseract-ocr poppler-utils

# macOS
brew install tesseract poppler
```

## Run Locally

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## Deploy on Render

1. Push to GitHub
2. New Web Service → connect repo
3. Build: `apt-get install -y tesseract-ocr poppler-utils && pip install -r requirements.txt`
4. Start: `gunicorn app:app`

Add `gunicorn` to requirements.txt for Render.

## Notes on SBI Card PDFs

- Works best with digital (text-based) PDFs — copy-paste works in these
- Scanned PDFs use OCR which may be slower and slightly less accurate
- Supports both formats: single Debit/Credit+Type column, or separate Debit/Credit columns
