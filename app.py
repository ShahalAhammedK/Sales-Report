import streamlit as st
import pandas as pd
import re
import os

DEFAULT_FILE_NAME = "sales_reports.xlsx"

# =========================================================
# This version is GENERIC: it works for any device's WhatsApp
# sales-report format (X300 Ultra, X300 Pro, V70FE, future
# devices, etc.) without needing separate code per device.
# It does this by:
#   1. Detecting the "Device" from the first line(s) of each report
#   2. Matching field LABELS loosely (lots of synonyms / spacing /
#      spelling variants), instead of hard-coding one device's wording
#   3. Splitting multi-report pastes using whichever anchor field
#      (a header line, "Sales Type:", or "Date:") repeats in the text
# =========================================================

# ---------- Field label patterns (synonym-tolerant) ----------
FIELD_PATTERNS = {
    "Sales Type":          r"Sales\s*Type\s*:[ \t\-]*([^\n]*)",
    "Date":                r"Date\s*:[ \t\-]*([^\n]*)",
    "Model Name":          r"Model\s*Name\s*:[ \t\-]*([^\n]*)",
    "VBA Name":            r"VBA\s*Name\s*:[ \t\-]*([^\n]*)",
    "Store":               r"Store(?:\s*Name)?\s*:[ \t\-]*([^\n]*)",
    "Customer Name":       r"(?:Customer|Coustmer)\s*Name\s*:[ \t\-]*([^\n]*)",
    "Color":               r"Colou?r\s*:[ \t\-]*([^\n]*)",
    "Variant/Option":      r"(?:Which\s*option(?:\s*\w+)?|Option|Storage\s*Variant|Variant|Package|RAM\s*\+?\s*ROM)\s*:[ \t\-]*([^\n]*)",
    "Nationality":         r"(?:Customer\s*)?National(?:ity)?\s*:[ \t\-]*([^\n]*)",
    "Occupation":          r"(?:Customer\s*)?Occupation\s*:[ \t\-]*([^\n]*)",
    "Previous Model Used": r"Previous\s*(?:which\s*)?model\s*Use(?:u)?d\s*:[ \t\-]*([^\n]*)",
    "Where did you hear":  r"Where\s*did\s*you\s*hear[^:\n]*:[ \t\-]*([^\n]*)",
}

COLUMNS = [
    "Device",
    "Sales Type",
    "Date",
    "VBA Name",
    "Store",
    "Customer Name",
    "Number",
    "Color",
    "Variant/Option",
    "Nationality",
    "Occupation",
    "Previous Model Used",
    "Where did you hear",
]


# ---------- Helpers ----------
def clean_value(value: str) -> str:
    value = value.strip()
    value = re.sub(r"[\-:\s]+$", "", value).strip()
    # treat dashes/punctuation-only values ("--", "—", ".") as empty
    if re.fullmatch(r"[\-\u2013\u2014\s\.]*", value):
        value = ""
    return value


# Matches any recognized field-label line ("Sales Type:", "Date:", "Number:",
# etc). Used to recognize where the field section starts, so anything ABOVE
# it (that isn't blank/a separator/a sender name) is treated as the device
# title — whatever it says (X300 Pro, X300 Ultra, V70 FE, or any future name).
FIELD_LABEL_RE = re.compile(
    r"^\s*\*?\s*(Sales\s*Type|Date|Model\s*Name|VBA\s*Name|Store(?:\s*Name)?|"
    r"(?:Customer|Coustmer)\s*Name|Colou?r|Which\s*option|Option|"
    r"Storage\s*Variant|Variant|Package|(?:Customer\s*)?National(?:ity)?|"
    r"(?:Customer\s*)?Occupation|Previous|Where\s*did\s*you\s*hear|"
    r"(?:Customer\s*|Contact\s*|Mobile\s*)?(?:Number|Nember|Nembr))[^\n:]*:",
    re.IGNORECASE,
)


# WhatsApp exports/forwards often prefix each message with
# "[9:55 pm, 29/06/2026] +971 56 651 5001: " before the actual report text.
# Strip that off so the sender's own phone number never gets mistaken for
# the customer's number, and so it doesn't pollute the device title.
WHATSAPP_PREFIX_RE = re.compile(r"^\s*\[[^\]\n]*\]\s*[^\n:]*:\s*")
ENVELOPE_DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")


def strip_whatsapp_prefix(text: str) -> str:
    return WHATSAPP_PREFIX_RE.sub("", text, count=1)


def extract_envelope_date(text: str) -> str:
    """If the WhatsApp timestamp bracket includes a full date with year
    (e.g. '[9:55 pm, 29/06/2026]'), use it as a fallback when the report
    body itself has no 'Date:' field. Brackets with only day/month and no
    year (e.g. '[08/05, 5:42 pm]') are skipped — too ambiguous to guess."""
    m = re.match(r"^\s*\[([^\]\n]*)\]", text)
    if not m:
        return ""
    date_match = ENVELOPE_DATE_RE.search(m.group(1))
    return date_match.group(0) if date_match else ""


def extract_device(text: str) -> str:
    """The device/model is whatever title is written above the field list —
    e.g. 'X300 Pro', 'vivo X300 Ultra reporting format', 'VIVO V70FE Sales
    Reporting', 'Vivo X300 pro'. We don't try to guess a model name from
    digits anywhere in the text (that misfires on phone numbers / storage
    sizes) — we just take the literal line(s) before the first field label."""
    lines = [l.strip(" *—-") for l in text.splitlines()]
    title_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if title_lines:
                break
            continue
        if re.match(r"^~", stripped):  # WhatsApp sender name line e.g. "~ A.b"
            continue
        if re.fullmatch(r"[\-=_—\s]{3,}", stripped):
            continue
        if FIELD_LABEL_RE.match(stripped):
            break
        title_lines.append(stripped)
        if len(title_lines) >= 2:
            break
    title = " ".join(title_lines) if title_lines else "Unknown"
    title = re.sub(r"^\d+\s*(?:st|nd|rd|th)[.,]?\s*", "", title, flags=re.IGNORECASE).strip()
    return title if title else "Unknown"


def extract_phone(text: str) -> str:
    # Try a UAE-style mobile number anywhere in the text first (label optional)
    phone_pattern = r'(\+?9715\d[\s\-]?\d{3}[\s\-]?\d{4}|0?5\d[\s\-]?\d{3}[\s\-]?\d{4})'
    phone = re.search(phone_pattern, text)
    if phone:
        return re.sub(r"[\s\-]", "", phone.group(0))

    # Fallback: look for a labelled number field even if format is odd
    m = re.search(
        r"(?:Customer\s*|Contact\s*|Mobile\s*)?Number\s*:?[ \t\-]*([^\n]*)",
        text, re.IGNORECASE,
    )
    if m:
        return clean_value(m.group(1))
    return ""


def normalize_sales_type(value: str) -> str:
    v = value.lower()
    if "pre" in v and ("book" in v or "order" in v):
        if "collect" in v:
            return "Pre-booking Collection" if "book" in v else "Pre-order Collection"
        return "Pre-booking" if "book" in v else "Pre-order"
    if "direct" in v:
        return "Direct Sale"
    return value


def normalize_date(value: str) -> str:
    if not value:
        return value
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"[\.\-]", "/", value)
    return value


# ---------- Extract Data from one report block ----------
def extract_data(text: str) -> dict:
    envelope_date = extract_envelope_date(text)
    text = text.replace("*", "")
    text = strip_whatsapp_prefix(text)
    text = re.sub(r"^\s*[—\-=_]{3,}\s*$", "", text, flags=re.MULTILINE)

    data = {"Device": extract_device(text)}

    for key, pattern in FIELD_PATTERNS.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = clean_value(match.group(1)) if match else ""

    # A "Model Name:" field (when present) is a more reliable device label
    # than a guessed title line like "Direct Sale".
    if data.get("Model Name"):
        data["Device"] = data["Model Name"]

    # Sales Type — fall back to scanning the whole message for cues if no label found
    if not data["Sales Type"]:
        tl = text.lower()
        if re.search(r"which\s*option\s*prebook", tl):
            data["Sales Type"] = "pre-booking"
        elif re.search(r"pre[\s\-]?booking\s*collection", tl):
            data["Sales Type"] = "pre-booking collection"
        elif re.search(r"pre[\s\-]?booking", tl):
            data["Sales Type"] = "pre-booking"
        elif re.search(r"pre[\s\-]?order\s*collection", tl):
            data["Sales Type"] = "pre-order collection"
        elif re.search(r"pre[\s\-]?order", tl):
            data["Sales Type"] = "pre-order"
        elif re.search(r"direct\s*sale", tl):
            data["Sales Type"] = "direct sale"
    data["Sales Type"] = normalize_sales_type(data["Sales Type"])

    data["Date"] = normalize_date(data["Date"]) or normalize_date(envelope_date)
    data["Number"] = extract_phone(text)

    return data


# Matches the START of an individual report within a blob of text: an
# optional ordinal marker ("1st,", "2nd*", "3rd,") followed by either a
# device/report header ("vivo V70 FE Sales Reporting", "X300 Ultra
# reporting format") or a standalone "Direct Sale" title line. This is what
# lets us split messages that bundle multiple numbered reports in one
# WhatsApp message ("1st... 2nd... 3rd...").
REPORT_START_RE = re.compile(
    r"^\s*\*?\s*"
    r"(?:\d+\s*(?:st|nd|rd|th)\b[.,]?\s*\*?\s*)?"
    r"(?:"
    r"(?:[A-Za-z]+\s+)?[A-Za-z]?\d{2,4}\s*\w*\s*(?:reporting\s*format|sales\s*reporting|reporting)\b"
    r"|Direct\s*Sale\b"
    r")",
    re.IGNORECASE | re.MULTILINE,
)


def _get_preceding_title(text: str, pos: int) -> str:
    """Walk backwards from `pos` collecting title/header lines (stopping at
    a blank line, a previous field-label line, or after 2 lines) so each
    split chunk keeps its own device title rather than the previous
    report's leftover text."""
    lines_before = text[:pos].splitlines()
    collected = []
    for line in reversed(lines_before):
        stripped = line.strip()
        if not stripped:
            break
        if FIELD_LABEL_RE.match(stripped):
            break
        collected.append(line)
        if len(collected) >= 2:
            break
    collected.reverse()
    return "\n".join(collected)


def _split_by_field_anchors(text: str):
    """Legacy fallback splitting, used when a text segment doesn't contain
    recognizable report-title lines (REPORT_START_RE). Anchors on whichever
    repeating field label is present: 'Sales Type:', a report header, or
    'Date:'."""

    sales_type_positions = [
        m.start() for m in re.finditer(r"^\s*\*?\s*Sales\s*Type\s*:",
                                        text, re.IGNORECASE | re.MULTILINE)
    ]
    if len(sales_type_positions) >= 2:
        chunks = []
        for i, start in enumerate(sales_type_positions):
            end = sales_type_positions[i + 1] if i + 1 < len(sales_type_positions) else len(text)
            title = _get_preceding_title(text, start)
            chunk = (title + "\n" if title else "") + text[start:end]
            chunks.append(chunk.strip())
        return chunks

    date_positions = [
        m.start() for m in re.finditer(r"^\s*\*?\s*Date\s*:",
                                        text, re.IGNORECASE | re.MULTILINE)
    ]
    if len(date_positions) >= 2:
        chunks = []
        for i, start in enumerate(date_positions):
            end = date_positions[i + 1] if i + 1 < len(date_positions) else len(text)
            title = _get_preceding_title(text, start)
            chunk = (title + "\n" if title else "") + text[start:end]
            chunks.append(chunk.strip())
        return chunks

    blocks = [b.strip() for b in re.split(r"\n\s*\n\s*\n+", text) if b.strip()]
    significant_blocks = [
        b for b in blocks
        if re.search(r"Customer\s*Name\s*:|Sales\s*Type\s*:", b, re.IGNORECASE)
    ]
    if len(significant_blocks) >= 2:
        return significant_blocks

    return [text.strip()] if text.strip() else []


def _split_segment(text: str):
    """Split a single text segment (already isolated to one WhatsApp
    envelope, or the whole paste if no envelopes were found) into
    individual reports."""
    matches = [m.start() for m in REPORT_START_RE.finditer(text)]
    if len(matches) >= 2:
        # The first match is only "redundant" (i.e. safe to merge into the
        # very start of the segment) when nothing but the envelope/title
        # precedes it. If report-1's own fields already appear before this
        # match, then it's a REAL boundary between two bundled reports
        # (e.g. report-1's title sat on the same line as the WhatsApp
        # envelope and so wasn't itself matchable) and must be kept.
        first_has_fields_before = bool(
            re.search(FIELD_LABEL_RE.pattern, text[:matches[0]], re.IGNORECASE | re.MULTILINE)
        )
        if first_has_fields_before:
            bounds = [0] + matches + [len(text)]
        else:
            bounds = [0] + matches[1:] + [len(text)]
        chunks = []
        for i in range(len(bounds) - 1):
            chunk = text[bounds[i]:bounds[i + 1]].strip()
            if chunk:
                chunks.append(chunk)
        return chunks
    return _split_by_field_anchors(text)


# ---------- Split into multiple reports ----------
def split_reports(text: str):
    """Split pasted text into individual reports, regardless of device type
    or how many reports are bundled into a single WhatsApp message."""

    # Split on WhatsApp message envelopes first, e.g.
    # "[08/05, 5:42 pm] +91 81025 78346: ..." or
    # "[9:55 pm, 29/06/2026] +971 56 651 5001: ...". This is the most
    # reliable top-level anchor for real chat exports/forwards.
    envelope_positions = [
        m.start() for m in re.finditer(r"^\s*\[[^\]\n]+\]\s*[^\n:]*:", text, re.MULTILINE)
    ]
    if len(envelope_positions) >= 2:
        segments = []
        for i, start in enumerate(envelope_positions):
            end = envelope_positions[i + 1] if i + 1 < len(envelope_positions) else len(text)
            segments.append(text[start:end])
    else:
        segments = [text]

    all_chunks = []
    for segment in segments:
        all_chunks.extend(_split_segment(segment))
    return all_chunks


# ---------- Save Excel ----------
def save_excel(df: pd.DataFrame, file_name: str):
    if os.path.exists(file_name):
        old = pd.read_excel(file_name, dtype=str)
        df = pd.concat([old, df], ignore_index=True)
    df.to_excel(file_name, index=False)


# ---------- UI ----------
st.set_page_config(page_title="Sales Report Collector", layout="wide")
st.title("📊 WhatsApp Sales Reporting Collector")
st.caption("Works across different devices / report formats — paste one or many reports at once.")

file_name = st.sidebar.text_input("Excel file name", value=DEFAULT_FILE_NAME)
if not file_name.endswith(".xlsx"):
    file_name += ".xlsx"

input_text = st.text_area(
    "Sales Reports",
    height=320,
    placeholder=(
        "Paste one or more WhatsApp sales reports here, e.g.:\n\n"
        "vivo X300 Ultra reporting format\n"
        "—————————\n"
        "Sales Type:- Direct sales\n"
        "Date:-29/06/2026\n"
        "VBA Name:- Muhammad Awais\n"
        "Store:-SDG Burjuman\n"
        "Customer Name:- Esha\n"
        "Number:- 0563413230\n"
        "Color:- Green\n"
        "Which option:- only Phone\n"
        "Nationality: india\n"
        "Occupation:-\n"
        "Previous which model Used:- Samsung\n"
        "Where did you hear about the X300?:-Promoter\n\n"
        "(You can paste a completely different device's report right "
        "after this one — it will be detected automatically.)"
    ),
)

col1, col2 = st.columns(2)

if col1.button("🔍 Extract Data"):
    if input_text.strip() == "":
        st.warning("Please paste at least one report")
    else:
        reports = split_reports(input_text)
        extracted = [extract_data(r) for r in reports]
        df = pd.DataFrame(extracted, columns=COLUMNS)

        # Drop rows that don't look like real reports
        df = df[df[["Date", "VBA Name", "Customer Name", "Number"]]
                .apply(lambda r: any(str(v).strip() for v in r), axis=1)]
        df = df.reset_index(drop=True)

        if df.empty:
            st.warning("No valid sales data found. Check your input format.")
        else:
            st.success(f"{len(df)} report(s) extracted")
            st.dataframe(df, use_container_width=True)
            st.session_state["data"] = df

if col2.button("💾 Save to Excel"):
    if "data" in st.session_state:
        save_excel(st.session_state["data"], file_name)
        st.success(f"Saved {len(st.session_state['data'])} row(s) to {file_name}")
    else:
        st.warning("Extract data first")

# ---------- Download ----------
if os.path.exists(file_name):
    with open(file_name, "rb") as f:
        st.download_button(
            "⬇ Download Excel",
            f,
            file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
