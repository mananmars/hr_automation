import streamlit as st
import pandas as pd
import re, io, zipfile
from docxtpl import DocxTemplate
import tempfile
import os
from docxtpl import DocxTemplate
import re


def extract_placeholders_from_docx(docx_file):
    """Extract unique Jinja-style placeholders {{ placeholder }} from a Word template"""
    doc = DocxTemplate(docx_file)
    placeholders = doc.get_undeclared_template_variables()
    # Deduplicate while preserving order
    return list(dict.fromkeys(placeholders))


def create_excel_from_placeholders(placeholders):
    """Create Excel with placeholder column headers"""
    df = pd.DataFrame(columns=placeholders)
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer

import io, os, zipfile, tempfile
import pandas as pd
from docxtpl import DocxTemplate

def generate_docs_from_excel(docx_template_file, df, name_placeholder="Candidate_Name"):
    """
    Fill Word template with Excel values and export as DOCX ZIP.
    Each file will be named as per the candidate's name (from name_placeholder column).
    
    :param docx_template_file: UploadedFile or path to .docx template
    :param df: DataFrame with filled placeholder values
    :param name_placeholder: column name in df to use for file naming
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for idx, row in df.iterrows():
            # Context for docxtpl
            context = {str(col).strip(): ("" if pd.isna(row[col]) else str(row[col])) for col in df.columns}

            # Render docx
            doc = DocxTemplate(docx_template_file)
            doc.render(context)

            # Candidate name for filename
            candidate_name = context.get(name_placeholder, f"Candidate_{idx+1}")
            safe_name = "".join(c for c in candidate_name if c.isalnum() or c in (" ", "_", "-")).rstrip()

            # Save docx temporarily
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, f"{safe_name}.docx")
                doc.save(docx_path)

                # Add DOCX to ZIP
                with open(docx_path, "rb") as f:
                    zipf.writestr(f"{safe_name}.docx", f.read())

    zip_buffer.seek(0)
    return zip_buffer

# -------------------- STREAMLIT APP --------------------
st.title("📄 HR Document Automation")

st.markdown("""
### Instructions:
1. Upload a **Word template (.docx)** containing placeholders like `{{ Candidate_Name }}` or `{{ DOJ }}`.  
2. Extract placeholders → download Excel template with these as columns.  
3. Fill Excel with values (one row per candidate).  
4. Upload the filled Excel → download documents as a ZIP.  
""")

# Step 1: Upload Docx Template
st.header("Step 1: Upload Word Template")
docx_file = st.file_uploader("Upload Word Template", type=["docx"])

placeholders = []

if docx_file:
    if st.button("Extract Placeholders"):
        placeholders = extract_placeholders_from_docx(docx_file)
        st.success(f"Extracted placeholders: {placeholders}")

        excel_file = create_excel_from_placeholders(placeholders)
        st.download_button("Download Excel Template", data=excel_file, file_name="placeholders.xlsx")

# Step 2: Upload Filled Excel
st.header("Step 2: Upload Filled Excel")
excel_file = st.file_uploader("Upload Excel with Data", type=["xlsx"])

if docx_file and excel_file:
    df = pd.read_excel(excel_file)
    if st.button("Generate PDF ZIP"):
        zip_buffer = generate_docs_from_excel(docx_file, df)
        st.download_button("Download All Documents(ZIP)", data=zip_buffer, file_name="documents.zip", mime="application/zip")
