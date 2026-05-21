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
    # in values like "Sales Type:- Direct"
    patterns = {
        "Sales Type":          r"Sales\s*Type\s*:[-\s]*(.*)",
        "Date":                r"Date\s*:[-\s]*(.*)",
        "VBA Name":            r"VBA\s*Name\s*:[-\s]*(.*)",
        "Store":               r"Store(?:\s*Name)?\s*:[-\s]*(.*)",
        "Customer Name":       r"Customer\s*Name\s*:[-\s]*(.*)",
        "Color":               r"Colou?r\s*:[-\s]*(.*)",
        "Which Option":        r"(?:Which\s*option|Option|Storage\s*Variant|Variant|Package)\s*:[-\s]*(.*)",
        "Nationality":         r"(?:Customer\s*)?Nationality\s*:[-\s]*(.*)",
        "Occupation":          r"(?:Customer\s*)?Occupation\s*:[-\s]*(.*)",
        "Previous Model Used": r"Previous\s*(?:which\s*)?model\s*Use(?:u)?d\s*:[-\s]*(.*)",
        "Where did you hear":  r"Where\s*did\s*you\s*hear[^:]*:[-\s]*(.*)",
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        value = match.group(1).strip() if match else ""
        # Stop at newline (keep only first line of the value)
        value = value.split("\n")[0].strip()
        # Remove trailing stars / dashes / colons left over
        value = re.sub(r"[\-:\s]+$", "", value).strip()
        data[key] = value

    # UAE phone extraction (works even if "Number:" label is missing)
    phone_pattern = r'(\+?9715\d[\s\-]?\d{3}[\s\-]?\d{4}|0?5\d[\s\-]?\d{3}[\s\-]?\d{4})'
    phone = re.search(phone_pattern, text)
    if phone:
        # Clean spaces/dashes in the number
        data["Number"] = re.sub(r"[\s\-]", "", phone.group(0))
    else:
        # Fallback: try generic "Number:" / "Contact Number:" label
        m = re.search(r"(?:Customer\s*)?(?:Contact\s*)?Number\s*:[-\s]*([^\n]+)", text, re.IGNORECASE)
        data["Number"] = m.group(1).strip() if m else ""

    # Normalize Sales Type
    st_val = data["Sales Type"].lower()
    if "pre" in st_val and ("book" in st_val or "order" in st_val):
        data["Sales Type"] = "Pre-booking Collection" if "collect" in st_val else "Pre-booking"
    elif "direct" in st_val:
        data["Sales Type"] = "Direct Sale"

    return data


# ---------- Split into multiple reports ----------
def split_reports(text):
    """Split pasted text into individual X300 Ultra reports."""
    # Strategy 1: split on the report header
    header_re = re.compile(
        r"(?=^\s*\*?\s*(?:vivo\s+)?x300\s*(?:ultra|pro)?.*reporting.*\*?\s*$)",
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
        "vivo X300 Ultra reporting format\n"
        "———————————————\n"
        "Sales Type:- Direct\n"
        "Date:- 21/05/2026\n"
        "VBA Name:- Arif\n"
        "Store:- SDG DCC\n"
        "Customer Name:- Dixon\n"
        "Number:- 0521668689\n"
        "Color:- Black\n"
        "Which option:- Full kit\n"
        "Nationality:- Lebanon\n"
        "Occupation:- Business\n"
        "Previous which model Used:- S23 Ultra\n"
        "Where did you hear about the X300?:- In Store"
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
