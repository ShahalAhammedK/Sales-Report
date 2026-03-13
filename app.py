import streamlit as st
import pandas as pd
import re
import os

# Set up the Page
st.set_page_config(page_title="Vivo V70 Sales Logger", page_icon="📱")
st.title("📱 Vivo V70 Sales Entry")
st.markdown("Paste the chat text below to automatically save it to Excel.")

# Regex Function
def parse_sales_data(text):
    patterns = {
        "Sales Type": r"Sales Type:\s*(.*)",
        "VBA Name": r"VBA Name:\*?(.*?)(?:\n|$)",
        "Store Name": r"Store Name:\*?(.*?)(?:\n|$)",
        "Customer Name": r"Customer Name:\*?(.*?)(?:\n|$)",
        "Customer Number": r"Customer Number:\s*(\d+)",
        "Storage Variant": r"Storage Variant:\s*(.*)",
        "Customer Nationality": r"Customer Nationality:\*?(.*?)(?:\n|$)",
        "Customer Occupation": r"Customer Occupation:\*?(.*?)(?:\n|$)"
    }
    data = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = match.group(1).strip("* ") if match else "N/A"
    return data

# Text Input Area
raw_input = st.text_area("Paste Chat Text Here:", height=250, placeholder="vivo V70 Sales Reporting...")

if st.button("Process & Save to Excel"):
    if raw_input.strip():
        new_entry = parse_sales_data(raw_input)
        filename = "Sales_Data.xlsx"
        
        # Load or Create DataFrame
        if os.path.exists(filename):
            df = pd.read_excel(filename)
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        else:
            df = pd.DataFrame([new_entry])
        
        # Save to Excel
        df.to_excel(filename, index=False)
        
        st.success("✅ Data Parsed and Saved!")
        st.table(pd.DataFrame([new_entry]).T) # Show what was added
        
        # Provide Download Link
        with open(filename, "rb") as file:
            st.download_button(
                label="📥 Download Updated Excel File",
                data=file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Please paste some text first!")