import streamlit as st
import pandas as pd
import re
import os

FILE_NAME = "vivo_sales.xlsx"

# ---------- Extract Data ----------
def extract_data(text):

    text = text.replace("*", "")  # remove *

    data = {}

    patterns = {
        "Sales Type": r"Sales Type:\s*(.*)",
        "VBA Name": r"VBA Name:\s*(.*)",
        "Store Name": r"Store Name:\s*(.*)",
        "Customer Name": r"Customer Name:\s*(.*)",
        "Storage Variant": r"Storage Variant:\s*(.*)",
        "Customer Nationality": r"Customer Nationality:\s*(.*)",
        "Customer Occupation": r"Customer Occupation:\s*(.*)"
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = match.group(1).strip() if match else ""

    # UAE phone extraction
    phone_pattern = r'(\+9715\d{8}|9715\d{8}|05\d{8})'
    phone = re.search(phone_pattern, text)
    data["Customer Number"] = phone.group(0) if phone else ""

    return data


# ---------- Save Excel ----------
def save_excel(df):

    if os.path.exists(FILE_NAME):
        old = pd.read_excel(FILE_NAME)
        df = pd.concat([old, df], ignore_index=True)

    df.to_excel(FILE_NAME, index=False)


# ---------- UI ----------
st.set_page_config(page_title="vivo Sales Collector", layout="wide")

st.title("📊 vivo V70 Sales Reporting Collector")

st.markdown("Paste **one or multiple reports** below")

input_text = st.text_area(
    "Sales Reports",
    height=300,
    placeholder="Paste reports here..."
)

col1, col2 = st.columns(2)

if col1.button("Extract Data"):

    if input_text.strip() == "":
        st.warning("Please paste a report")
    else:

        reports = input_text.split("vivo V70 FE")

        extracted = []

        for r in reports:
            if r.strip():
                r = "vivo V70 FE" + r
                extracted.append(extract_data(r))

        df = pd.DataFrame(extracted)

        st.success(f"{len(df)} report(s) extracted")

        st.dataframe(df, use_container_width=True)

        st.session_state["data"] = df


if col2.button("Save to Excel"):

    if "data" in st.session_state:
        save_excel(st.session_state["data"])
        st.success("Saved to Excel")
    else:
        st.warning("Extract data first")


# ---------- Download ----------
if os.path.exists(FILE_NAME):

    with open(FILE_NAME, "rb") as f:
        st.download_button(
            "⬇ Download Excel",
            f,
            FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
