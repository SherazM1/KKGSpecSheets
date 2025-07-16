import streamlit as st
import pandas as pd
from pdftoexcel import extract_pdf_data, make_excel_file_from_data  # <- Import your new functions!
from pdftoexcel import field_order, field_aliases 
import os



if os.path.exists("kkglogo.png"):
    st.image("kkglogo.png", width=180)

st.set_page_config(
    page_title="Kendal King Spec Sheet App",
    layout="centered",
)

st.title("Spec Sheet PDF to Excel Tool")
st.markdown(
    "Upload one or more spec sheet PDFs. See a preview of the extracted data, then download as Excel."
)

# Upload PDF(s)
uploaded_files = st.file_uploader(
    "Upload PDF file(s)", type=["pdf"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Please upload at least one PDF file to continue.")
    st.stop()

# Process all PDFs
all_rows = []
for file in uploaded_files:
    data = extract_pdf_data(file, field_order, field_aliases)
    all_rows.extend(data)

# Convert to DataFrame for preview
df = pd.DataFrame(all_rows, columns=field_order)
st.subheader("Extracted Data Preview")
st.dataframe(df, height=400)

# Name for Excel file
excel_name = st.text_input("Excel file name", value="spec_sheets.xlsx")

# Button to download Excel
if st.button("Download Excel"):
    excel_buffer = make_excel_file_from_data(all_rows, field_order, file_name=excel_name)
    st.download_button(
        label="⬇️ Download Excel",
        data=excel_buffer,
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
