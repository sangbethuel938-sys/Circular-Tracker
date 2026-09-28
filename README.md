# EATTA Circular Register

Local Streamlit application that:

1. uploads a circular as PDF/JPG/PNG;
2. reads text from the circular (native PDF text, or OCR when Tesseract is installed);
3. pre-fills Date, To, From and Subject/RE where possible;
4. reserves a unique serial number in Google Sheets, e.g. `EATTA/CIR/2026/001`;
5. stamps the number in green in the upper-right area above **Circular to Members**;
6. saves the numbered PDF locally under `Circulars/<YEAR>/`;
7. completes the Google Sheets register row;
8. keeps a local SQLite + Excel backup.

## 1. Install Python

Install Python 3.10 or newer from python.org. During installation tick **Add Python to PATH**.

## 2. Create the Google Sheet

Create a blank Google Sheet named **EATTA Circular Register**.

Open **Extensions > Apps Script** and replace the default code with the contents of `appscript.gs` in this repository. Save the project.

The script automatically creates a tab named **Circular Register** and its headers on first use.

## 3. Add the shared secret in Apps Script

In Apps Script open **Project Settings > Script properties** and add:

- Property: `APP_SHARED_SECRET`
- Value: a long random secret known only to you.

Example format (do not use this exact value):

`eatta-circular-2026-a-long-random-private-key`

## 4. Deploy Apps Script as a Web App

In Apps Script:

1. Select **Deploy > New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**.
4. Who has access: **Anyone**.
5. Deploy and authorize.
6. Copy the Web app URL ending in `/exec`.

The endpoint requires the shared secret for every register operation.

## 5. Configure Streamlit secrets

Copy:

`.streamlit/secrets.toml.example`

to:

`.streamlit/secrets.toml`

Then fill in:

```toml
APPS_SCRIPT_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"
APPS_SCRIPT_SECRET = "the-same-secret-you-added-to-apps-script"
```

**Never commit `secrets.toml` to GitHub.** It is already excluded by `.gitignore`.

## 6. Install and run on Windows

Double-click:

`install_and_run.bat`

For later use, double-click:

`run_app.bat`

Or run manually:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit normally opens at `http://localhost:8501`.

## 7. OCR for scanned/image-only circulars

Searchable PDFs are read without extra software. For image-only scans, install **Tesseract OCR for Windows**.

After installation, make sure `tesseract.exe` is available on PATH. If it is installed in a non-standard location, add its installation folder to Windows PATH.

If Tesseract is not installed, the app still works; you can type/correct the Date, To, From and Subject fields manually.

## 8. Git / GitHub

From the project folder:

```bash
git init
git add .
git commit -m "Initial EATTA circular register"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

Do not add `.streamlit/secrets.toml` to the repository.

## File structure

```text
EATTA_Circular_Register/
├── app.py
├── appscript.gs
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   └── secrets.toml.example
├── Circulars/       # numbered PDFs created here
├── data/            # SQLite local backup
└── Reports/         # Excel local backup
```

## Numbering and audit protection

The Apps Script uses `LockService` when reserving a number. That makes the serial generator safe when more than one person uses the app.

If a number is reserved but the PDF save fails, the number remains **RESERVED** in Google Sheets and is not reused. This is intentional for audit integrity.
