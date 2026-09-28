from __future__ import annotations

import io
import os
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import pandas as pd
import requests
import streamlit as st
from PIL import Image

try:
    import pytesseract
except Exception:
    pytesseract = None


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
CIRCULARS_DIR = APP_DIR / "Circulars"
REPORTS_DIR = APP_DIR / "Reports"
DB_PATH = DATA_DIR / "circular_register.db"
EXCEL_PATH = REPORTS_DIR / "Circular_Register.xlsx"

for p in (DATA_DIR, CIRCULARS_DIR, REPORTS_DIR):
    p.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="EATTA Circular Register", page_icon="📄", layout="wide")

BLUE = (0.0, 0.35, 0.75)  # PDF RGB blue for circular number


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS circulars (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                circular_number TEXT UNIQUE NOT NULL,
                serial_no INTEGER,
                date TEXT,
                subject TEXT,
                recipient TEXT,
                sender TEXT,
                status TEXT,
                file_name TEXT,
                local_file_path TEXT,
                date_recorded TEXT,
                remarks TEXT,
                extracted_text TEXT
            )
            """
        )
        con.commit()


def save_local_record(rec: dict[str, Any]) -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """
            INSERT INTO circulars (
                circular_number, serial_no, date, subject, recipient, sender,
                status, file_name, local_file_path, date_recorded, remarks, extracted_text
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(circular_number) DO UPDATE SET
                date=excluded.date,
                subject=excluded.subject,
                recipient=excluded.recipient,
                sender=excluded.sender,
                status=excluded.status,
                file_name=excluded.file_name,
                local_file_path=excluded.local_file_path,
                date_recorded=excluded.date_recorded,
                remarks=excluded.remarks,
                extracted_text=excluded.extracted_text
            """,
            (
                rec["circular_number"], rec.get("serial_no"), rec.get("date", ""),
                rec.get("subject", ""), rec.get("to", ""), rec.get("from", ""),
                rec.get("status", ""), rec.get("file_name", ""),
                rec.get("local_file_path", ""), rec.get("date_recorded", ""),
                rec.get("remarks", ""), rec.get("extracted_text", ""),
            ),
        )
        con.commit()
    export_excel()


def export_excel() -> None:
    with sqlite3.connect(DB_PATH) as con:
        df = pd.read_sql_query(
            """
            SELECT serial_no AS 'Serial No.', circular_number AS 'Circular Number',
                   date AS 'Date', subject AS 'Subject', recipient AS 'To', sender AS 'From',
                   status AS 'Status', file_name AS 'File Name',
                   local_file_path AS 'Local File Path', date_recorded AS 'Date Recorded',
                   remarks AS 'Remarks'
            FROM circulars ORDER BY id DESC
            """,
            con,
        )
    df.to_excel(EXCEL_PATH, index=False)


def local_records() -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as con:
        return pd.read_sql_query(
            """
            SELECT serial_no AS 'Serial No.', circular_number AS 'Circular Number',
                   date AS 'Date', subject AS 'Subject', recipient AS 'To', sender AS 'From',
                   status AS 'Status', file_name AS 'File Name',
                   local_file_path AS 'Local File Path', date_recorded AS 'Date Recorded',
                   remarks AS 'Remarks'
            FROM circulars ORDER BY id DESC
            """,
            con,
        )


def get_secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default)).strip()
    except Exception:
        return os.environ.get(name, default).strip()


def api_call(action: str, **kwargs) -> dict[str, Any]:
    url = get_secret("APPS_SCRIPT_URL")
    secret = get_secret("APPS_SCRIPT_SECRET")
    if not url or not secret:
        raise RuntimeError("Google Sheets connection is not configured. Add APPS_SCRIPT_URL and APPS_SCRIPT_SECRET in Streamlit Secrets.")
    payload = {"action": action, "secret": secret, **kwargs}
    r = requests.post(url, json=payload, timeout=30)
    r.raise_for_status()
    data = r.json()
    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Google Sheets request failed."))
    return data


def read_pdf_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text("text") for page in doc).strip()
    doc.close()
    return text


def ocr_image(image: Image.Image) -> str:
    if pytesseract is None:
        return ""
    try:
        return pytesseract.image_to_string(image)
    except Exception:
        return ""


def ocr_pdf(pdf_bytes: bytes) -> str:
    if pytesseract is None:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        chunks = []
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            chunks.append(ocr_image(img))
        doc.close()
        return "\n".join(chunks).strip()
    except Exception:
        return ""


def extract_text(uploaded) -> str:
    data = uploaded.getvalue()
    ext = Path(uploaded.name).suffix.lower()
    if ext == ".pdf":
        text = read_pdf_text(data)
        if len(re.sub(r"\s+", "", text)) >= 80:
            return text
        return ocr_pdf(data)
    img = Image.open(io.BytesIO(data)).convert("RGB")
    return ocr_image(img)


def clean_line(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip(" :-\t")


def parse_fields(text: str) -> dict[str, str]:
    """Extract common EATTA circular fields from PDF/OCR text.

    Handles:
    - To: All Members
    - Multi-line member association recipient lists
    - From: Secretariat
    - Split-line labels produced by PDF extraction/OCR
    """
    lines = [clean_line(x) for x in text.splitlines() if clean_line(x)]

    labels = {"to", "from", "copy to", "date", "re", "subject"}

    def labelled_value(label: str, max_follow: int = 1) -> str:
        wanted = label.lower()
        for i, line in enumerate(lines):
            # Same-line form: From: Secretariat / Date: 26th August 2021
            m = re.match(rf"(?i)^{re.escape(label)}\s*:\s*(.*)$", line)
            if m:
                first = clean_line(m.group(1))
                if first:
                    return first

                vals = []
                for candidate in lines[i + 1 : i + 1 + max_follow]:
                    low = candidate.lower().rstrip(":")
                    if low in labels or re.match(
                        r"(?i)^(To|From|Copy to|Date|RE|Subject)\s*:", candidate
                    ):
                        break
                    vals.append(candidate)
                return "; ".join(vals)

            # Split-line form: "From" then "Secretariat"
            if line.lower().rstrip(":") == wanted:
                vals = []
                for candidate in lines[i + 1 : i + 1 + max_follow]:
                    low = candidate.lower().rstrip(":")
                    if low in labels or re.match(
                        r"(?i)^(To|From|Copy to|Date|RE|Subject)\s*:", candidate
                    ):
                        break
                    vals.append(candidate)
                return "; ".join(vals)

        return ""

    raw_date = labelled_value("Date", 1)
    sender = labelled_value("From", 2)
    subject = labelled_value("RE", 2) or labelled_value("Subject", 2)

    # Recipient handling:
    # 1. If the circular explicitly says "All Members", preserve that exact meaning.
    # 2. Otherwise collect the multi-line To block until the next main label.
    recipient = ""

    if re.search(r"(?i)\ball\s+members\b", text):
        recipient = "All Members"
    else:
        for i, line in enumerate(lines):
            m = re.match(r"(?i)^To\s*:\s*(.*)$", line)

            if m or line.lower().rstrip(":") == "to":
                recipients = []

                if m:
                    first = clean_line(m.group(1))
                    if first:
                        if re.search(r"(?i)\ball\s+members\b", first):
                            recipient = "All Members"
                            break
                        recipients.append(first)

                for candidate in lines[i + 1 : i + 10]:
                    if re.match(
                        r"(?i)^(From|Copy to|Date|RE|Subject)\s*:?(?:\s|$)",
                        candidate,
                    ):
                        break

                    if re.search(r"(?i)\ball\s+members\b", candidate):
                        recipient = "All Members"
                        break

                    recipients.append(candidate)

                if not recipient:
                    recipient = "; ".join(
                        dict.fromkeys(x for x in recipients if x)
                    )
                break

    return {
        "date": raw_date,
        "subject": subject,
        "to": recipient,
        "from": sender,
    }


def uploaded_to_pdf(uploaded) -> bytes:
    data = uploaded.getvalue()
    ext = Path(uploaded.name).suffix.lower()
    if ext == ".pdf":
        return data
    img = Image.open(io.BytesIO(data)).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PDF", resolution=150.0)
    return buf.getvalue()


def stamp_pdf(pdf_bytes: bytes, circular_number: str) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    # Designed around the supplied EATTA letterhead: below contact/header block,
    # above "Circular to Members", right aligned.
    # Uses page-relative coordinates so it also behaves well on A4/Letter scans.
    x_right = page.rect.width - 48
    x_left = max(page.rect.width * 0.55, x_right - 250)
    y_top = page.rect.height * 0.145
    y_bottom = y_top + 18
    rect = fitz.Rect(x_left, y_top, x_right, y_bottom)

    page.insert_textbox(
        rect,
        circular_number,
        fontsize=9.5,
        fontname="helv",
        color=BLUE,
        align=fitz.TEXT_ALIGN_RIGHT,
        overlay=True,
    )
    out = doc.tobytes(garbage=4, deflate=True)
    doc.close()
    return out


def safe_name(number: str) -> str:
    return number.replace("/", "_") + ".pdf"


def save_pdf(pdf_bytes: bytes, circular_number: str) -> Path:
    year_match = re.search(r"/(\d{4})/", circular_number)
    year = year_match.group(1) if year_match else str(datetime.now().year)
    folder = CIRCULARS_DIR / year
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / safe_name(circular_number)
    path.write_bytes(pdf_bytes)
    return path


init_db()

st.title("EATTA Circular Register")
st.caption("Upload a scanned circular → read details → reserve serial in Google Sheets → stamp green number → save locally → update register.")

with st.sidebar:
    st.header("EATTA Circular Register")

    configured = bool(
        get_secret("APPS_SCRIPT_URL")
        and get_secret("APPS_SCRIPT_SECRET")
    )

    st.subheader("Google Sheets")
    if configured:
        st.success("Configured")
    else:
        st.warning("Not configured")

    if st.button(
        "Test Google Sheets connection",
        use_container_width=True,
        disabled=not configured,
    ):
        try:
            api_call("health")
            st.success("Connection is working.")
        except Exception as e:
            st.error(f"Connection failed: {e}")

    st.divider()

    st.subheader("Circular numbering")
    st.caption("Automatic reference format")
    st.code("EATTA/CIR/YYYY/001", language=None)

    st.subheader("Storage")
    st.caption("Numbered circular PDFs")
    st.code(str(CIRCULARS_DIR), language=None)

    st.caption("Local register backup")
    st.code(str(DB_PATH), language=None)

    st.caption("Excel backup")
    st.code(str(EXCEL_PATH), language=None)

    st.divider()

    st.subheader("Workflow")
    st.markdown(
        """
        1. Upload circular
        2. Check extracted details
        3. Reserve circular number
        4. Stamp and save PDF
        5. Update Google Sheet
        """
    )

create_tab, register_tab = st.tabs(["Create Circular", "Register"])

with create_tab:
    uploaded = st.file_uploader("Upload scanned circular", type=["pdf", "png", "jpg", "jpeg"])

    if uploaded:
        upload_key = f"{uploaded.name}:{uploaded.size}"
        if st.session_state.get("upload_key") != upload_key:
            with st.spinner("Reading circular..."):
                text = extract_text(uploaded)
                parsed = parse_fields(text)
            st.session_state.upload_key = upload_key
            st.session_state.extracted_text = text
            st.session_state.parsed = parsed
            st.session_state.pop("reserved_number", None)
            st.session_state.pop("reserved_serial", None)

        parsed = st.session_state.get("parsed", {})
        text = st.session_state.get("extracted_text", "")

        if text:
            st.success("Text was read from the circular. Check the details below before saving.")
        else:
            st.warning("No readable text was detected. Enter the fields manually. For image-only scans, install Tesseract OCR as described in README.md.")

        c1, c2 = st.columns(2)
        with c1:
            circular_date = st.text_input("Circular date", value=parsed.get("date", ""), placeholder="e.g. 28th September 2026")
            sender = st.text_input("From", value=parsed.get("from", ""))
        with c2:
            recipient = st.text_area("To", value=parsed.get("to", ""), height=90)
            status = st.selectbox("Status", ["Issued", "Draft"])

        subject = st.text_area("Subject / RE", value=parsed.get("subject", ""), height=90)
        remarks = st.text_input("Remarks", value="")

        if not st.session_state.get("reserved_number"):
            if st.button("1. Reserve next circular number", type="primary", use_container_width=True):
                if not configured:
                    st.error("Configure the Apps Script URL and shared secret first.")
                else:
                    try:
                        result = api_call("reserve", year=date.today().year)
                        st.session_state.reserved_number = result["circular_number"]
                        st.session_state.reserved_serial = int(result["serial_no"])
                        st.rerun()
                    except Exception as e:
                        st.error(f"Could not reserve serial: {e}")
        else:
            number = st.session_state.reserved_number
            st.info(f"Reserved circular number: **{number}**")

            if st.button("2. Stamp, save PDF and update Google Sheet", type="primary", use_container_width=True):
                try:
                    pdf_bytes = uploaded_to_pdf(uploaded)
                    stamped = stamp_pdf(pdf_bytes, number)
                    saved_path = save_pdf(stamped, number)
                    recorded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    record = {
                        "serial_no": st.session_state.reserved_serial,
                        "circular_number": number,
                        "date": circular_date,
                        "subject": subject,
                        "to": recipient,
                        "from": sender,
                        "status": status,
                        "file_name": saved_path.name,
                        "local_file_path": str(saved_path),
                        "date_recorded": recorded_at,
                        "remarks": remarks,
                        "extracted_text": text,
                    }

                    # Local backup first: the numbered PDF and metadata are never lost.
                    save_local_record(record)

                    # Complete the reserved row in Google Sheets.
                    api_call(
                        "complete",
                        circular_number=number,
                        date=circular_date,
                        subject=subject,
                        to=recipient,
                        **{"from": sender},
                        status=status,
                        file_name=saved_path.name,
                        local_file_path=str(saved_path),
                        remarks=remarks,
                    )

                    st.success(f"Saved successfully as {number}")
                    st.download_button(
                        "Download numbered circular",
                        data=stamped,
                        file_name=saved_path.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
                    st.session_state.pop("reserved_number", None)
                    st.session_state.pop("reserved_serial", None)
                except Exception as e:
                    st.error(f"Save failed: {e}")
                    st.warning("If a serial was already reserved, it remains reserved and will not be reused. This protects the audit trail.")

        with st.expander("Extracted text"):
            st.text_area("OCR / PDF text", value=text, height=260, disabled=True, label_visibility="collapsed")

with register_tab:
    st.subheader("Circular Register")
    source = st.radio("View records from", ["Google Sheets", "Local backup"], horizontal=True)

    if source == "Google Sheets" and configured:
        try:
            data = api_call("list").get("records", [])
            df = pd.DataFrame(data)
            if not df.empty:
                df = df.rename(columns={
                    "serial_no": "Serial No.", "circular_number": "Circular Number",
                    "date": "Date", "subject": "Subject", "to": "To", "from": "From",
                    "status": "Status", "file_name": "File Name",
                    "local_file_path": "Local File Path", "date_recorded": "Date Recorded",
                    "remarks": "Remarks"
                })
        except Exception as e:
            st.error(f"Could not load Google Sheet: {e}")
            df = local_records()
    else:
        df = local_records()

    if df.empty:
        st.info("No circulars recorded yet.")
    else:
        search = st.text_input("Search register", placeholder="Circular number, subject, recipient...")
        shown = df.copy()
        if search:
            mask = shown.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            shown = shown[mask]
        st.dataframe(shown, use_container_width=True, hide_index=True)

        if EXCEL_PATH.exists():
            st.download_button(
                "Download local Excel backup",
                data=EXCEL_PATH.read_bytes(),
                file_name="Circular_Register.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
