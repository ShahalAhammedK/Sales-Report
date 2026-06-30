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
    r"^\s*\*?\s*(Sales\s*Type|Date|VBA\s*Name|Store(?:\s*Name)?|"
    r"(?:Customer|Coustmer)\s*Name|Colou?r|Which\s*option|Option|"
    r"Storage\s*Variant|Variant|Package|(?:Customer\s*)?National(?:ity)?|"
    r"(?:Customer\s*)?Occupation|Previous|Where\s*did\s*you\s*hear|"
    r"(?:Customer\s*|Contact\s*|Mobile\s*)?Number)[^\n:]*:",
    re.IGNORECASE,
)


# WhatsApp exports/forwards often prefix each message with
# "[9:55 pm, 29/06/2026] +971 56 651 5001: " before the actual report text.
# Strip that off so the sender's own phone number never gets mistaken for
# the customer's number, and so it doesn't pollute the device title.
WHATSAPP_PREFIX_RE = re.compile(r"^\s*\[[^\]\n]*\]\s*[^\n:]*:\s*")


def strip_whatsapp_prefix(text: str) -> str:
    return WHATSAPP_PREFIX_RE.sub("", text, count=1)


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
    return " ".join(title_lines) if title_lines else "Unknown"


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
    text = text.replace("*", "")
    text = strip_whatsapp_prefix(text)
    text = re.sub(r"^\s*[—\-=_]{3,}\s*$", "", text, flags=re.MULTILINE)

    data = {"Device": extract_device(text)}

    for key, pattern in FIELD_PATTERNS.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = clean_value(match.group(1)) if match else ""

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

    data["Date"] = normalize_date(data["Date"])
    data["Number"] = extract_phone(text)

    return data


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


# ---------- Split into multiple reports ----------
def split_reports(text: str):
    """Split pasted text into individual reports, regardless of device type."""

    # Strategy 1: split on repeated "Sales Type:" lines. This is the most
    # reliable anchor since nearly every report format includes it,
    # regardless of whether it also has a recognizable header line.
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

    # Strategy 2: split on generic report-header lines, e.g.
    # "vivo X300 Ultra reporting format", "VIVO V70FE Sales Reporting",
    # "X300 Pro Sales Reporting", etc.
    header_re = re.compile(
        r"(?=^\s*\*?\s*(?:[A-Za-z]+\s+)?[A-Za-z]?\d{2,4}\s*\w*\s*"
        r"(?:reporting\s*format|sales\s*reporting|reporting)\b.*$)",
        re.IGNORECASE | re.MULTILINE,
    )
    parts = [p.strip() for p in header_re.split(text) if p.strip()]
    if len(parts) >= 2:
        return parts

    # Strategy 3: split on repeated "Date:" lines (reports without Sales Type)
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

    # Strategy 4: split on blank-line-separated blocks (last resort, only if
    # there are at least 2 blocks that each contain a customer name / number)
    blocks = [b.strip() for b in re.split(r"\n\s*\n\s*\n+", text) if b.strip()]
    if len(blocks) >= 2:
        return blocks

    return [text.strip()] if text.strip() else []


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
