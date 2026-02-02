import re
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader
from io import BytesIO

st.set_page_config(page_title="PDF Table Extractor", layout="wide")

st.title("📄 Invoice PDF Table Extractor")
st.write("Upload one or multiple PDFs and extract invoice table data into a single Excel file.")

# Regex pattern for table rows
pattern = r"^(\d+)\s+(\d{2}/\d{2}/\d{4})\s+([A-Z]{2})\s+(\d+)\s+([\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+([\d,]+\.\d+)"

# Upload PDFs
uploaded_files = st.file_uploader(
    "Upload Invoice PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        st.subheader(f"Processing: {file.name}")

        reader = PdfReader(file)
        text = ""

        # Extract full text from PDF
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

        # Extract matches
        matches = re.findall(pattern, text, re.MULTILINE)

        if matches:
            df = pd.DataFrame(matches, columns=[
                "Line",
                "Invoice Date",
                "Province",
                "Invoice Number",
                "Gross Extended",
                "Discount",
                "Subtotal Before Tax"
            ])

            # Convert numeric columns
            for col in ["Gross Extended", "Discount", "Subtotal Before Tax"]:
                df[col] = df[col].str.replace(",", "").astype(float)

            df["Line"] = df["Line"].astype(int)

            # Add source file info
            df["Source PDF"] = file.name

            all_data.append(df)

            st.success(f"✅ Extracted {len(df)} rows from {file.name}")
            st.dataframe(df)

        else:
            st.warning(f"⚠ No invoice table data found in {file.name}")

    # Combine all extracted data
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)

        st.markdown("---")
        st.subheader("✅ Combined Output Table")
        st.dataframe(final_df)

        # Create Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Invoice Data", index=False)

        st.download_button(
            label="📥 Download the extracted data",
            data=output.getvalue(),
            file_name="CanadianTire_Extracted_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("No valid table data extracted from uploaded PDFs.")
