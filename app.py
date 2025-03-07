import streamlit as st
import pandas as pd
import re
import os

def extract_narration_details(narration):
    if pd.isna(narration):
        return [""] * 7
    if narration.startswith("CAM/"):
        match = re.search(r"(\d{2}[-./]\d{2}[-./]\d{2,4})", narration)
        txn_date = match.group(1) if match else ""
        narration = narration.replace(txn_date, "__DATE__")
    else:
        txn_date = ""
    if narration.startswith("IMPS Chg"):
        return [narration] + [""] * 6
    split_data = re.split(r"[/-]", narration)
    split_data = [item.strip() for item in split_data if item.strip()]
    if "__DATE__" in split_data:
        split_data[split_data.index("__DATE__")] = txn_date
    while len(split_data) < 7:
        split_data.append("")
    return split_data[:7]

st.title("Narration Data Extractor")
st.write("Upload an Excel file to extract structured data from the 'Narration' column.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    if 'Narration' in df.columns:
        df_narration_split = df[['Narration']].copy()
        column_names = [f"Field {i+1}" for i in range(7)]
        df_narration_split[column_names] = df['Narration'].apply(lambda x: pd.Series(extract_narration_details(x)))
        
        st.write("### Extracted Data Preview")
        st.dataframe(df_narration_split.head())
        
        output_file = "Extracted_Narration.xlsx"
        df_narration_split.to_excel(output_file, index=False)
        
        st.download_button(label="Download Extracted Data", data=open(output_file, "rb"), file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("The uploaded file does not contain a 'Narration' column.")
