from __future__ import annotations

import io
import os
import re
import sqlite3
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import pandas as pd
import requests
import streamlit as st
from PIL import Image, ImageOps, ImageFilter

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
                extracted_text TEXT,
                document_type TEXT DEFAULT 'Original',
                original_circular TEXT,
                amendment_no INTEGER,
                amendment_reason TEXT
            )
            """
        )

        existing = {
            row[1]
            for row in con.execute("PRAGMA table_info(circulars)").fetchall()
        }
        migrations = {
            "document_type": "ALTER TABLE circulars ADD COLUMN document_type TEXT DEFAULT 'Original'",
            "original_circular": "ALTER TABLE circulars ADD COLUMN original_circular TEXT",
            "amendment_no": "ALTER TABLE circulars ADD COLUMN amendment_no INTEGER",
            "amendment_reason": "ALTER TABLE circulars ADD COLUMN amendment_reason TEXT",
        }

        for column, statement in migrations.items():
            if column not in existing:
                con.execute(statement)

        con.commit()


def save_local_record(rec: dict[str, Any]) -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """
            INSERT INTO circulars (
                circular_number, serial_no, date, subject, recipient, sender,
                status, file_name, local_file_path, date_recorded, remarks, extracted_text,
                document_type, original_circular, amendment_no, amendment_reason
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                extracted_text=excluded.extracted_text,
                document_type=excluded.document_type,
                original_circular=excluded.original_circular,
                amendment_no=excluded.amendment_no,
                amendment_reason=excluded.amendment_reason
            """,
            (
                rec["circular_number"], rec.get("serial_no"), rec.get("date", ""),
                rec.get("subject", ""), rec.get("to", ""), rec.get("from", ""),
                rec.get("status", ""), rec.get("file_name", ""),
                rec.get("local_file_path", ""), rec.get("date_recorded", ""),
                rec.get("remarks", ""), rec.get("extracted_text", ""),
                rec.get("document_type", "Original"),
                rec.get("original_circular", ""),
                rec.get("amendment_no"),
                rec.get("amendment_reason", ""),
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


def configure_tesseract() -> tuple[bool, str]:
    """Locate and configure the Tesseract executable.

    Works on Streamlit Cloud/Linux and common Windows installs.
    Returns (available, diagnostic_message).
    """
    if pytesseract is None:
        return False, "pytesseract Python package is not installed."

    candidates = [
        shutil.which("tesseract"),
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    ]

    for candidate in candidates:
        if candidate and Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return True, f"Tesseract found at {candidate}"

    return False, "Tesseract executable was not found."


def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    """Improve scanned-document readability before OCR."""
    img = image.convert("L")
    img = ImageOps.autocontrast(img)

    # Upscale smaller scans so Tesseract has enough pixel detail.
    if img.width < 1800:
        scale = max(1.0, 1800 / max(img.width, 1))
        img = img.resize(
            (int(img.width * scale), int(img.height * scale)),
            Image.Resampling.LANCZOS,
        )

    img = img.filter(ImageFilter.SHARPEN)
    return img


def ocr_image(image: Image.Image) -> str:
    available, _ = configure_tesseract()
    if not available:
        return ""

    try:
        prepared = preprocess_for_ocr(image)

        # PSM 6 is reliable for letter/circular layouts with blocks of text.
        text = pytesseract.image_to_string(
            prepared,
            lang="eng",
            config="--oem 3 --psm 6",
        )
        return text.strip()
    except Exception:
        return ""


def ocr_pdf(pdf_bytes: bytes) -> str:
    available, _ = configure_tesseract()
    if not available:
        return ""

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        chunks = []

        # 3x rendering gives OCR a much clearer image than 2x on scans.
        for page in doc:
            pix = page.get_pixmap(
                matrix=fitz.Matrix(3, 3),
                alpha=False,
            )
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            page_text = ocr_image(img)
            if page_text:
                chunks.append(page_text)

        doc.close()
        return "\n\n".join(chunks).strip()

    except Exception:
        return ""


def extract_text(uploaded) -> str:
    data = uploaded.getvalue()
    ext = Path(uploaded.name).suffix.lower()

    if ext == ".pdf":
        native_text = read_pdf_text(data)

        # If the PDF already contains a useful text layer, use it.
        if len(re.sub(r"\s+", "", native_text)) >= 80:
            return native_text

        # Otherwise treat it as an image-only / scanned PDF.
        ocr_text = ocr_pdf(data)
        return ocr_text or native_text

    img = Image.open(io.BytesIO(data)).convert("RGB")
    return ocr_image(img)


def clean_line(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip(" :-\t")


def parse_fields(text: str) -> dict[str, str]:
    """Extract fields from both labelled and unlabelled circular/letter formats.

    Supports:
    - EATTA circulars with To / From / Date / RE / Subject labels
    - "All Members"
    - Formal letters with address blocks followed by Dear Sir/Madam
    - Subject headings written in uppercase without a Subject/RE label
    - Sender details inferred from Yours faithfully/sincerely signature blocks
    - Dates written as 7th September 2026, 07/09/2026, 2026-09-07, etc.

    The app still presents all extracted values for user review before saving.
    """
    if not text:
        return {"date": "", "subject": "", "to": "", "from": ""}

    raw_lines = [x.strip() for x in text.replace("\r", "\n").splitlines()]
    lines = [clean_line(x) for x in raw_lines if clean_line(x)]

    label_re = re.compile(
        r"(?i)^(to|from|copy\s*to|date|re|subject)\s*[:\-]?\s*(.*)$"
    )

    parsed_by_label: dict[str, list[str]] = {
        "to": [],
        "from": [],
        "copy to": [],
        "date": [],
        "re": [],
        "subject": [],
    }

    current_label = None

    for line in lines:
        m = label_re.match(line)

        if m:
            label = re.sub(r"\s+", " ", m.group(1).lower()).strip()
            value = clean_line(m.group(2))
            current_label = label

            if value:
                parsed_by_label.setdefault(label, []).append(value)
            continue

        # Multi-line recipient blocks are common in EATTA circulars.
        if current_label in {"to", "copy to"}:
            if re.match(
                r"(?i)^(from|date|re|subject|copy\s*to)\b",
                line,
            ):
                current_label = None
            else:
                parsed_by_label[current_label].append(line)
                continue

        # Single-line or short continuation values.
        elif current_label in {"from", "date", "re", "subject"}:
            if not parsed_by_label[current_label]:
                parsed_by_label[current_label].append(line)
            current_label = None

    # ---------------------------------------------------------
    # DATE
    # ---------------------------------------------------------
    date_parts = parsed_by_label.get("date", [])
    raw_date = clean_line(date_parts[0]) if date_parts else ""

    date_patterns = [
        r"\b\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b",
        r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b",
    ]

    # If no labelled date, only inspect the TOP portion of the document first.
    # This avoids mistaking event dates in the body for the circular date.
    if not raw_date:
        top_lines = lines[:10]

        for top_line in top_lines:
            found = None
            for pattern in date_patterns:
                m = re.search(pattern, top_line, flags=re.I)
                if m:
                    found = m.group(0)
                    break
            if found:
                raw_date = clean_line(found)
                break

    # ---------------------------------------------------------
    # TO / RECIPIENT
    # ---------------------------------------------------------
    recipient = ""

    if re.search(r"(?i)\ball\s+members\b", text):
        recipient = "All Members"
    else:
        to_parts = parsed_by_label.get("to", [])

        if to_parts:
            cleaned_parts = []

            for part in to_parts:
                part = clean_line(part)
                if not part:
                    continue

                if re.match(
                    r"(?i)^(from|copy\s*to|date|re|subject)\b",
                    part,
                ):
                    break

                cleaned_parts.append(part)

            recipient = "; ".join(dict.fromkeys(cleaned_parts))

    # Formal-letter fallback:
    # infer recipient/address block immediately before "Dear Sir/Madam".
    if not recipient:
        salutation_index = None

        for i, line in enumerate(lines):
            if re.match(
                r"(?i)^dear\s+(sir|madam|sir/madam|sirs|team|members?)\b",
                line,
            ):
                salutation_index = i
                break

        if salutation_index is not None and salutation_index > 0:
            start_idx = max(0, salutation_index - 6)
            candidates = lines[start_idx:salutation_index]

            address_parts = []

            for part in candidates:
                if re.match(r"(?i)^ref(?:erence)?\s*[:\-]", part):
                    continue

                if any(re.search(p, part, flags=re.I) for p in date_patterns):
                    continue

                if re.match(
                    r"(?i)^(east african tea trade association|tea board of kenya|directors?)$",
                    part,
                ):
                    continue

                # Avoid obvious letterhead/contact lines.
                if re.search(
                    r"(?i)(tel:|telephone|mobile|email|e-mail|www\.|p\.?o\.?\s*box\s+\d{5,})",
                    part,
                ):
                    # PO Box can belong to recipient, so retain short address lines.
                    if not re.match(r"(?i)^p\.?o\.?\s*box", part):
                        continue

                address_parts.append(part)

            # Prefer the last few lines nearest the salutation.
            if address_parts:
                recipient = "; ".join(address_parts[-5:])

    # ---------------------------------------------------------
    # SUBJECT / RE
    # ---------------------------------------------------------
    subject_parts = parsed_by_label.get("re", []) or parsed_by_label.get("subject", [])
    subject = clean_line(" ".join(subject_parts)) if subject_parts else ""

    if not subject:
        # Same-line unlabelled OCR variants.
        m = re.search(
            r"(?im)^(?:re|subject)\s*[:\-]?\s+(.+)$",
            text,
        )
        if m:
            subject = clean_line(m.group(1))

    if not subject:
        # Look for a title-like uppercase line near the start of the main body.
        # This handles letters such as "COLLECTION OF TEA SAMPLES FOR ANALYSIS".
        heading_candidates = []

        for idx, line in enumerate(lines[:30]):
            if len(line) < 8 or len(line) > 180:
                continue

            if re.match(
                r"(?i)^(dear\s+|yours\s+|east african tea trade association|tea board of kenya|directors?)",
                line,
            ):
                continue

            letters = re.sub(r"[^A-Za-z]", "", line)

            if len(letters) < 8:
                continue

            uppercase_ratio = sum(c.isupper() for c in letters) / max(len(letters), 1)

            # Prefer uppercase headings and lines immediately after salutation.
            if uppercase_ratio >= 0.72:
                score = len(line)

                if idx > 0 and re.match(r"(?i)^dear\s+", lines[idx - 1]):
                    score += 100

                heading_candidates.append((score, line))

        if heading_candidates:
            subject = max(heading_candidates, key=lambda x: x[0])[1]

    # ---------------------------------------------------------
    # FROM / SENDER
    # ---------------------------------------------------------
    sender_parts = parsed_by_label.get("from", [])
    sender = clean_line(sender_parts[0]) if sender_parts else ""

    if not sender:
        # Locate closing/signature block.
        closing_index = None

        for i, line in enumerate(lines):
            if re.match(
                r"(?i)^yours\s+(faithfully|sincerely|truly)\b",
                line,
            ):
                closing_index = i
                break

        if closing_index is not None:
            signature_lines = lines[closing_index + 1 : closing_index + 8]

            useful = []

            for part in signature_lines:
                part = clean_line(part)

                if not part:
                    continue

                # Ignore obvious signature/scribble OCR noise.
                alpha = re.sub(r"[^A-Za-z]", "", part)
                if len(alpha) < 3:
                    continue

                # Stop when footer/contact section begins.
                if re.search(
                    r"(?i)(p\.?o\.?\s*box|telephone|mobile|email|e-mail|www\.|tea house|nyerere avenue)",
                    part,
                ):
                    break

                useful.append(part)

            if useful:
                # Prefer organization/name/title lines.
                role_terms = (
                    "director",
                    "managing director",
                    "chief executive officer",
                    "ceo",
                    "secretariat",
                    "manager",
                    "chairman",
                    "secretary",
                )

                selected = []

                for part in useful:
                    low = part.lower()

                    if (
                        any(term in low for term in role_terms)
                        or "association" in low
                        or "board" in low
                        or part.isupper()
                    ):
                        selected.append(part)

                if not selected:
                    selected = useful[:3]

                sender = "; ".join(dict.fromkeys(selected[:4]))

    # Specific EATTA fallback if OCR sees Secretariat close to a From marker.
    if not sender and re.search(r"(?i)\bsecretariat\b", text):
        m = re.search(
            r"(?is)\bfrom\s*[:\-]?\s*(?:\n|\s)+([^\n]{1,80})",
            text,
        )
        if m:
            sender = clean_line(m.group(1))

    return {
        "date": raw_date,
        "subject": subject,
        "to": recipient,
        "from": sender,
    }


def parse_circular_date(value: str) -> date:
    """Convert extracted circular dates to a standard Python date.

    Handles examples such as:
    - 26th August 2021
    - 26 August 2021
    - 26/08/2021
    - 2021-08-26

    Falls back to today's date when OCR cannot confidently parse the value.
    """
    if not value:
        return date.today()

    text = clean_line(value)

    # Remove ordinal suffixes: 1st, 2nd, 3rd, 4th...
    text = re.sub(r"(?i)(\d{1,2})(st|nd|rd|th)\b", r"\1", text)

    formats = [
        "%d %B %Y",
        "%d %b %Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    # Try to find a date-like part inside longer OCR text.
    patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b",
        r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            candidate = match.group(0)
            for fmt in formats:
                try:
                    return datetime.strptime(candidate, fmt).date()
                except ValueError:
                    pass

    return date.today()


def format_circular_date(value: date) -> str:
    """Store/display circular dates in standard DD/MM/YYYY format."""
    return value.strftime("%d/%m/%Y")


def base_circular_number(number: str) -> str:
    """Return original circular reference without /AMn suffix."""
    return re.sub(r"/AM\d+$", "", str(number or "").strip(), flags=re.I)


def available_original_circulars() -> list[str]:
    """Return original circulars available for amendment selection."""
    options: list[str] = []

    try:
        rows = api_call("list").get("records", [])
        for row in rows:
            number = str(row.get("circular_number", "")).strip()
            document_type = str(row.get("document_type", "")).strip().lower()

            if number and "/AM" not in number.upper() and document_type != "amendment":
                options.append(number)
    except Exception:
        try:
            df = local_records()
            if "Circular Number" in df.columns:
                for value in df["Circular Number"].tolist():
                    number = str(value).strip()
                    if number and "/AM" not in number.upper():
                        options.append(number)
        except Exception:
            pass

    return sorted(set(options), reverse=True)


def uploaded_to_pdf(uploaded) -> bytes:
    data = uploaded.getvalue()
    ext = Path(uploaded.name).suffix.lower()
    if ext == ".pdf":
        return data
    img = Image.open(io.BytesIO(data)).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PDF", resolution=150.0)
    return buf.getvalue()


def detect_green_header_line_y(page) -> float | None:
    """Detect the long green horizontal line in the upper part of the page.

    Returns the line position in PDF points, or None if no reliable line is found.
    This lets the circular number sit just below the actual letterhead line even
    when a scan is slightly shifted, resized, or cropped.
    """
    try:
        # Only inspect the upper 35% of page; render at 2x for reliable detection.
        clip = fitz.Rect(
            0,
            0,
            page.rect.width,
            page.rect.height * 0.35,
        )
        matrix = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=matrix, clip=clip, alpha=False)

        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        width, height = img.size
        pixels = img.load()

        best_row = None
        best_count = 0

        # Ignore the extreme top edge and scan for a green-dominant horizontal row.
        for y in range(max(2, int(height * 0.05)), height - 2):
            count = 0

            # Sample every second pixel for speed.
            for x in range(0, width, 2):
                r, g, b = pixels[x, y]

                # Broad EATTA-green detector. Designed to tolerate scan variation.
                if (
                    g >= 65
                    and g > r * 1.18
                    and g > b * 1.12
                    and (g - r) >= 18
                ):
                    count += 1

            if count > best_count:
                best_count = count
                best_row = y

        # Require a reasonably long line. Text alone should not meet this threshold.
        sampled_width = max(1, width // 2)
        if best_row is None or best_count < sampled_width * 0.22:
            return None

        # Convert rendered pixel Y back into PDF-point coordinates.
        return clip.y0 + (best_row / 2.0)

    except Exception:
        return None



def find_empty_stamp_position(page, circular_number: str, fontsize: float = 9.5) -> tuple[float, float]:
    """Find a clean empty place near the top of the page for the circular number.

    Strategy:
    - Render the top portion of the page as an image.
    - Test several candidate positions across the upper area.
    - Score each candidate by how much visible content already exists there.
    - Prefer cleaner areas, while still preferring the top-right/top area.
    """
    text_width = fitz.get_text_length(
        circular_number,
        fontname="helv",
        fontsize=fontsize,
    )

    box_width = text_width + 12
    box_height = fontsize + 10

    # Search only the top band of the page.
    top_band_height = min(page.rect.height * 0.28, 170)
    clip = fitz.Rect(0, 0, page.rect.width, top_band_height)

    try:
        scale = 2.0
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        width, height = img.size
        pixels = img.load()

        detected_line_y = None
        try:
            detected_line_y = detect_green_header_line_y(page)
        except Exception:
            detected_line_y = None

        # Candidate X positions (prefer right side first, but allow center/left if cleaner).
        x_candidates = []
        right_margin = 36
        left_margin = 24

        x_right = page.rect.width - right_margin - box_width
        x_center = max(left_margin, (page.rect.width - box_width) / 2)
        x_left = left_margin

        # Add more dense candidates across the upper band
        fractions = [0.72, 0.58, 0.44, 0.30, 0.16]
        for frac in fractions:
            x_candidates.append(max(left_margin, page.rect.width * frac - box_width / 2))

        # Deduplicate while preserving preference order: right, center, custom, left.
        ordered_x = []
        for x in [x_right, x_center, *x_candidates, x_left]:
            x = min(max(left_margin, x), max(left_margin, page.rect.width - box_width - 12))
            if not any(abs(x - existing) < 3 for existing in ordered_x):
                ordered_x.append(x)

        # Candidate Y positions.
        y_candidates = [22, 30, 38, 46, 54, 64, 74, 86, 98, 110]

        # If the green line is detected in a sensible header zone, prioritize just below it.
        if detected_line_y is not None:
            if page.rect.height * 0.05 <= detected_line_y <= page.rect.height * 0.22:
                preferred_y = detected_line_y + 8
                y_candidates = [preferred_y] + y_candidates

        best = None
        best_score = None

        for y_pdf in y_candidates:
            if y_pdf + box_height > top_band_height:
                continue

            for x_pdf in ordered_x:
                # Convert the candidate box to rendered-image coordinates.
                x0 = int(max(0, x_pdf * scale))
                y0 = int(max(0, y_pdf * scale))
                x1 = int(min(width, (x_pdf + box_width) * scale))
                y1 = int(min(height, (y_pdf + box_height) * scale))

                if x1 <= x0 or y1 <= y0:
                    continue

                total = 0
                occupied = 0

                # Count darker pixels as occupied content.
                for yy in range(y0, y1):
                    for xx in range(x0, x1):
                        total += 1
                        if pixels[xx, yy] < 225:
                            occupied += 1

                occupancy_ratio = occupied / max(total, 1)

                # Preference penalties:
                # - small penalty for positions lower down
                # - small penalty for moving left away from the right/top area
                right_penalty = (page.rect.width - (x_pdf + box_width)) / max(page.rect.width, 1) * 0.02
                down_penalty = (y_pdf / max(top_band_height, 1)) * 0.03

                score = occupancy_ratio + right_penalty + down_penalty

                if best_score is None or score < best_score:
                    best_score = score
                    best = (x_pdf, y_pdf)

        if best is not None:
            return best

    except Exception:
        pass

    # Safe fallback
    fallback_x = max(24, page.rect.width - 36 - box_width)
    fallback_y = 30
    return fallback_x, fallback_y


def stamp_pdf(pdf_bytes: bytes, circular_number: str) -> bytes:
    """Stamp the circular number horizontally in a clean empty place near the top.

    This version is page-rotation aware. Some scanned PDFs are internally stored
    with 90/180/270-degree page rotation, which previously caused a horizontal
    stamp to appear vertically down the side.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    fontsize = 9.5

    # find_empty_stamp_position works in the visible / rotated page coordinate space.
    visible_x, visible_y = find_empty_stamp_position(
        page,
        circular_number=circular_number,
        fontsize=fontsize,
    )

    # Convert the desired visible position back to the PDF's unrotated coordinate
    # system, then rotate the inserted text by the page's own rotation so it
    # appears horizontal to the reader.
    visible_point = fitz.Point(visible_x, visible_y)
    insert_point = visible_point * page.derotation_matrix

    page.insert_text(
        insert_point,
        circular_number,
        fontsize=fontsize,
        fontname="helv",
        color=BLUE,
        rotate=page.rotation,
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

    st.subheader("OCR")
    ocr_ready, ocr_message = configure_tesseract()
    if ocr_ready:
        st.success("OCR ready")
    else:
        st.warning("OCR unavailable")
        st.caption(ocr_message)

    st.divider()

    st.subheader("Circular numbering")
    st.caption("Automatic reference format")
    st.code("EATTA/CIR/YYYY/001\nEATTA/CIR/YYYY/001/AM1", language=None)

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
            st.success(
                "Text was read from the circular and the fields below were auto-filled. "
                "Review and correct anything before saving."
            )
        else:
            ocr_ready, ocr_message = configure_tesseract()
            if ocr_ready:
                st.warning(
                    "No readable text was detected from this scan. "
                    "Try a clearer/higher-resolution scan or enter the fields manually."
                )
            else:
                st.error(
                    "OCR is not available on this deployment. "
                    f"{ocr_message} Check packages.txt and requirements.txt."
                )

        st.subheader("Document classification")
        document_type = st.radio(
            "Document type",
            ["Original", "Amendment"],
            horizontal=True,
            help="Choose Amendment when this circular changes an earlier circular.",
        )

        original_circular = ""
        amendment_reason = ""

        if document_type == "Amendment":
            original_options = available_original_circulars()

            if original_options:
                original_circular = st.selectbox(
                    "Original circular being amended",
                    options=original_options,
                    help="The amendment reference will be generated as /AM1, /AM2, etc.",
                )
            else:
                original_circular = st.text_input(
                    "Original circular being amended",
                    placeholder="EATTA/CIR/2026/001",
                )

            amendment_reason = st.text_area(
                "Reason / summary of amendment",
                placeholder="Briefly state what is being amended.",
                height=80,
            )

        c1, c2 = st.columns(2)
        with c1:
            detected_date = parse_circular_date(parsed.get("date", ""))
            selected_date = st.date_input(
                "Circular date",
                value=detected_date,
                format="DD/MM/YYYY",
                help="Select or correct the circular date. It will be saved as DD/MM/YYYY.",
            )
            circular_date = format_circular_date(selected_date)

            sender = st.text_input(
                "From",
                value=parsed.get("from", ""),
            )

        with c2:
            recipient = st.text_area(
                "To",
                value=parsed.get("to", ""),
                height=90,
            )
            status = st.selectbox(
                "Status",
                ["Issued", "Draft"],
            )

        subject = st.text_area(
            "Subject / RE",
            value=parsed.get("subject", ""),
            height=90,
        )
        remarks = st.text_input("Remarks", value="")

        if not st.session_state.get("reserved_number"):
            if st.button("1. Reserve next circular number", type="primary", use_container_width=True):
                if not configured:
                    st.error("Configure the Apps Script URL and shared secret first.")
                else:
                    try:
                        if document_type == "Amendment":
                            if not original_circular:
                                raise RuntimeError("Select or enter the original circular being amended.")

                            result = api_call(
                                "reserve_amendment",
                                original_circular=base_circular_number(original_circular),
                            )
                        else:
                            result = api_call(
                                "reserve",
                                year=selected_date.year,
                            )

                        st.session_state.reserved_number = result["circular_number"]
                        st.session_state.reserved_serial = int(result.get("serial_no") or 0)
                        st.session_state.reserved_document_type = document_type
                        st.session_state.reserved_original_circular = result.get(
                            "original_circular",
                            base_circular_number(original_circular) if original_circular else "",
                        )
                        st.session_state.reserved_amendment_no = result.get("amendment_no")
                        st.session_state.reserved_amendment_reason = amendment_reason
                        st.rerun()
                    except Exception as e:
                        st.error(f"Could not reserve serial: {e}")
        else:
            number = st.session_state.reserved_number
            reserved_type = st.session_state.get("reserved_document_type", "Original")
            reserved_original = st.session_state.get("reserved_original_circular", "")
            reserved_am_no = st.session_state.get("reserved_amendment_no")
            reserved_reason = st.session_state.get("reserved_amendment_reason", "")

            st.info(f"Reserved circular number: **{number}**")

            if reserved_type == "Amendment":
                st.caption(
                    f"Amendment {reserved_am_no} to {reserved_original}"
                    + (f" — {reserved_reason}" if reserved_reason else "")
                )

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
                        "document_type": reserved_type,
                        "original_circular": reserved_original if reserved_type == "Amendment" else number,
                        "amendment_no": reserved_am_no if reserved_type == "Amendment" else 0,
                        "amendment_reason": reserved_reason if reserved_type == "Amendment" else "",
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
                        document_type=reserved_type,
                        original_circular=reserved_original if reserved_type == "Amendment" else number,
                        amendment_no=reserved_am_no if reserved_type == "Amendment" else 0,
                        amendment_reason=reserved_reason if reserved_type == "Amendment" else "",
                    )

                    st.success(f"Saved successfully as {number}")
                    st.download_button(
                        "Download numbered circular",
                        data=stamped,
                        file_name=saved_path.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
                    for key in (
                        "reserved_number",
                        "reserved_serial",
                        "reserved_document_type",
                        "reserved_original_circular",
                        "reserved_amendment_no",
                        "reserved_amendment_reason",
                    ):
                        st.session_state.pop(key, None)
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
                    "remarks": "Remarks",
                    "document_type": "Document Type",
                    "original_circular": "Original Circular",
                    "amendment_no": "Amendment No.",
                    "amendment_reason": "Amendment Reason",
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

        st.subheader("Amendment trace")
        originals = sorted(
            {
                base_circular_number(v)
                for v in shown["Circular Number"].astype(str).tolist()
                if str(v).strip()
            },
            reverse=True,
        )

        if originals:
            trace_base = st.selectbox(
                "Select circular to trace",
                originals,
                key="trace_base_circular",
            )

            trace = shown[
                shown["Circular Number"]
                .astype(str)
                .str.startswith(trace_base, na=False)
            ].copy()

            trace_columns = [
                c for c in [
                    "Circular Number",
                    "Document Type",
                    "Amendment No.",
                    "Date",
                    "Subject",
                    "Amendment Reason",
                    "Status",
                ]
                if c in trace.columns
            ]

            if trace_columns:
                st.dataframe(
                    trace[trace_columns],
                    use_container_width=True,
                    hide_index=True,
                )

        if EXCEL_PATH.exists():
            st.download_button(
                "Download local Excel backup",
                data=EXCEL_PATH.read_bytes(),
                file_name="Circular_Register.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
