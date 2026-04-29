# PDF Unlocker

A web app to remove passwords from multiple PDFs at once.

## Features
- Upload up to 50 PDFs at a time
- Enter one password → tries against all files
- Files that unlock are ready to download instantly
- Files that fail show in a retry panel — enter a different password and retry
- Download all unlocked files as a ZIP (or single PDF if only one)

## Run Locally

```bash
pip install flask pypdf
python app.py
```

Then open http://localhost:5000

## Deploy on Render

1. Push this folder to a GitHub repo
2. New Web Service on Render → connect repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`

Add `gunicorn` to requirements.txt for Render deployment.
