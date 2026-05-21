import streamlit as st
import pandas as pd
import re
import os

FILE_NAME = "x300_ultra_sales.xlsx"

# ---------- Extract Data ----------
def extract_data(text):
    # Clean WhatsApp markup
    text = text.replace("*", "")
    text = re.sub(r"^\s*[—\-=_]{3,}\s*$", "", text, flags=re.MULTILINE)

    data = {}

    # Patterns — each handles ":" or ":-" separator and ignores leading dashes
    # in values like "Sales Type:- Direct".
    # Note: we use [ \t\-]* (not [-\s]*) after the colon so the separator
    # cannot greedily consume a newline if the value is empty (e.g. "Number:- --").
    patterns = {
        "Sales Type":          r"Sales\s*Type\s*:[ \t\-]*([^\n]*)",
        "Date":                r"Date\s*:[ \t\-]*([^\n]*)",
        "VBA Name":            r"VBA\s*Name\s*:[ \t\-]*([^\n]*)",
        "Store":               r"Store(?:\s*Name)?\s*:[ \t\-]*([^\n]*)",
        "Customer Name":       r"(?:Customer|Coustmer)\s*Name\s*:[ \t\-]*([^\n]*)",
        "Color":               r"Colou?r\s*:[ \t\-]*([^\n]*)",
        "Which Option":        r"(?:Which\s*option(?:\s*\w+)?|Option|Storage\s*Variant|Variant|Package)\s*:[ \t\-]*([^\n]*)",
        "Nationality":         r"(?:Customer\s*)?Nationality\s*:[ \t\-]*([^\n]*)",
        "Occupation":          r"(?:Customer\s*)?Occupation\s*:[ \t\-]*([^\n]*)",
        "Previous Model Used": r"Previous\s*(?:which\s*)?model\s*Use(?:u)?d\s*:[ \t\-]*([^\n]*)",
        "Where did you hear":  r"Where\s*did\s*you\s*hear[^:\n]*:[ \t\-]*([^\n]*)",
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        value = match.group(1).strip() if match else ""
        value = re.sub(r"[\-:\s]+$", "", value).strip()
        # If the value is just dashes/punctuation (e.g. "--", "-", "—"), treat as empty
        if re.fullmatch(r"[\-\u2013\u2014\s\.]*", value):
            value = ""
        data[key] = value

    # UAE phone extraction (works even if "Number:" label is missing or "--")
    phone_pattern = r'(\+?9715\d[\s\-]?\d{3}[\s\-]?\d{4}|0?5\d[\s\-]?\d{3}[\s\-]?\d{4})'
    phone = re.search(phone_pattern, text)
    if phone:
        data["Number"] = re.sub(r"[\s\-]", "", phone.group(0))
    else:
        # Fallback: "Number:" / "Customer Number:" / "Contact Number:" label.
        # Use [ \t\-]* (not [-\s]*) so we don't accidentally consume the newline
        # when the value is just "--" or empty.
        m = re.search(
            r"(?:Customer\s*|Contact\s*)?Number\s*:[ \t\-]*([^\n]*)",
            text, re.IGNORECASE,
        )
        if m:
            val = m.group(1).strip()
            val = re.sub(r"[\-:\s]+$", "", val).strip()
            # Treat dashes-only ("--", "—") as empty
            if re.fullmatch(r"[\-\u2013\u2014\s\.]*", val):
                val = ""
            data["Number"] = val
        else:
            data["Number"] = ""

    # Sales Type — first try the labelled value, then fall back to scanning the
    # WHOLE message for cues. Strong cues: "prebooking" / "pre-booking" header,
    # "Which option prebooked:" label (only pre-booking reports use this),
    # "PRE ORDER" / "Direct Sale" mentions, etc.
    st_val = data["Sales Type"].lower()
    if not st_val:
        text_lower = text.lower()
        # "Which option prebooked" is a 100% pre-booking signal
        if re.search(r"which\s*option\s*prebook", text_lower):
            st_val = "pre-booking"
        elif re.search(r"pre[\s\-]?booking\s*collection", text_lower):
            st_val = "pre-booking collection"
        elif re.search(r"pre[\s\-]?booking", text_lower):
            st_val = "pre-booking"
        elif re.search(r"pre[\s\-]?order\s*collection", text_lower):
            st_val = "pre-order collection"
        elif re.search(r"pre[\s\-]?order", text_lower):
            st_val = "pre-order"
        elif re.search(r"direct\s*sale", text_lower):
            st_val = "direct sale"

    if "pre" in st_val and ("book" in st_val or "order" in st_val):
        data["Sales Type"] = "Pre-booking Collection" if "collect" in st_val else "Pre-booking"
    elif "direct" in st_val:
        data["Sales Type"] = "Direct Sale"

    # Date normalization — "12.06.2026" → "12/06/2026", "12-06-2026" → "12/06/2026"
    if data["Date"]:
        data["Date"] = re.sub(r"[\.\-]", "/", data["Date"]).strip()

    return data


# ---------- Split into multiple reports ----------
def split_reports(text):
    """Split pasted text into individual X300 Ultra reports."""
    # Strategy 1: split on any X300 report header
    # Matches: "vivo X300 Ultra reporting format", "vivo X300 Ultra prebooking",
    # "X300 Pro Sales Reporting", "vivo X300 Ultra Pre-booking", etc.
    header_re = re.compile(
        r"(?=^\s*\*?\s*(?:vivo\s+)?x300\s*(?:ultra|pro|fe)?\s*"
        r"(?:reporting|pre[\s\-]?booking|pre[\s\-]?order|sales)\b.*$)",
        re.IGNORECASE | re.MULTILINE,
    )
    parts = header_re.split(text)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) >= 2:
        return parts

    # Strategy 2: split on repeated "Sales Type:" lines
    sales_type_positions = [
        m.start() for m in re.finditer(r"^\s*\*?\s*Sales\s*Type\s*:",
                                       text, re.IGNORECASE | re.MULTILINE)
    ]
    if len(sales_type_positions) >= 2:
        chunks = []
        for i, start in enumerate(sales_type_positions):
            end = sales_type_positions[i + 1] if i + 1 < len(sales_type_positions) else len(text)
            chunks.append(text[start:end].strip())
        return chunks

    # Strategy 3: split on repeated "Date:" lines (for reports without Sales Type label)
    date_positions = [
        m.start() for m in re.finditer(r"^\s*\*?\s*Date\s*:",
                                       text, re.IGNORECASE | re.MULTILINE)
    ]
    if len(date_positions) >= 2:
        chunks = []
        for i, start in enumerate(date_positions):
            end = date_positions[i + 1] if i + 1 < len(date_positions) else len(text)
            # Include the line before (might be the header)
            line_start = text.rfind("\n", 0, start) + 1
            prev_line_start = text.rfind("\n", 0, line_start - 1) + 1
            chunks.append(text[prev_line_start:end].strip())
        return chunks

    # Single report
    return [text.strip()] if text.strip() else []


# ---------- Save Excel ----------
def save_excel(df):
    if os.path.exists(FILE_NAME):
        old = pd.read_excel(FILE_NAME, dtype=str)
        df = pd.concat([old, df], ignore_index=True)
    df.to_excel(FILE_NAME, index=False)


# ---------- Column order ----------
COLUMNS = [
    "Sales Type",
    "Date",
    "VBA Name",
    "Store",
    "Customer Name",
    "Number",
    "Color",
    "Which Option",
    "Nationality",
    "Occupation",
    "Previous Model Used",
    "Where did you hear",
]


# ---------- UI ----------
st.set_page_config(page_title="X300 Ultra Sales Collector", layout="wide")
st.title("📊 vivo X300 Ultra Sales Reporting Collector")
st.markdown("Paste **one or multiple reports** below — then click Extract.")

input_text = st.text_area(
    "Sales Reports",
    height=320,
    placeholder=(
        "vivo X300 Ultra prebooking\n"
        "———————————————\n"
        "Date:- 12.06.2026\n"
        "VBA Name:- Mohammed sahil\n"
        "Store:- Vivo Al tamam store\n"
        "Customer Name:- Nishad yusuf\n"
        "Customer Number:- 0501234567\n"
        "Color:- black\n"
        "Which option prebooked:- only phone\n"
        "Nationality:- Indian\n"
        "Occupation:- Business\n"
        "Previous which model Used:- S26 ULTRA\n"
        "Where did you hear about the X300?: Social media"
    ),
)

col1, col2 = st.columns(2)

if col1.button("🔍 Extract Data"):
    if input_text.strip() == "":
        st.warning("Please paste a report")
    else:
        reports = split_reports(input_text)
        extracted = [extract_data(r) for r in reports]
        df = pd.DataFrame(extracted, columns=COLUMNS)
        # Drop empty rows
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
        save_excel(st.session_state["data"])
        st.success(f"Saved {len(st.session_state['data'])} row(s) to {FILE_NAME}")
    else:
        st.warning("Extract data first")

# ---------- Download ----------
if os.path.exists(FILE_NAME):
    with open(FILE_NAME, "rb") as f:
        st.download_button(
            "⬇ Download Excel",
            f,
            FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
