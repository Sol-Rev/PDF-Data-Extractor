import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# --------- Utility functions ---------
def clean_value(val: str):
    if val is None:
        return None
    return val.replace("-", "").strip()

def extract_from_pdf(file, filename):
    results = {
        "File": filename,
        "Invoice Number": None,
        "Invoice Date": None,
        "ORDER #": None,
        "INVOICE TOTAL": None,
        "TOTAL": None,
        "TOTAL GST": None,
    }

    try:
        with pdfplumber.open(file) as pdf:
            total_pages = len(pdf.pages)
            last_page_index = total_pages - 1
            page = pdf.pages[last_page_index]
            text = page.extract_text(x_tolerance=1, y_tolerance=1, layout=True).split("\n")

            # --- Invoice Number & Date ---
            invoice_index = [i+2 for i, x in enumerate(text) if "INVOICE NUMBER" in x]
            if invoice_index:
                parts = text[invoice_index[0]].split()
                results["Invoice Number"] = parts[-3] if len(parts) >= 3 else None
                results["Invoice Date"] = parts[-2] if len(parts) >= 2 else None

            # --- ORDER # ---
            order_index = [i+1 for i, x in enumerate(text) if "ORDER #" in x]
            if order_index:
                parts = text[order_index[0]].split()
                results["ORDER #"] = parts[-5] if len(parts) >= 5 else None

            # --- INVOICE TOTAL ---
            invoice_total_index = [i for i, x in enumerate(text) if "INVOICE TOTAL" in x]
            if invoice_total_index:
                parts = text[invoice_total_index[0]].split()
                invoice_total = parts[-1] if len(parts) >= 2 else None
                results["INVOICE TOTAL"] = clean_value(invoice_total)

            # --- TOTAL & GST ---
            total_gst_index = [i for i, x in enumerate(text) if "TOTAL -" in x]
            total_sum, gst_sum = 0.0, 0.0

            for idx in total_gst_index:
                parts = text[idx].split()
                nums = []
                for p in parts:
                    if not re.search(r"\d", p):
                        continue
                    cleaned = re.sub(r"[^0-9.\-]", "", p).strip()
                    if cleaned.endswith("-"):
                        cleaned = cleaned[:-1]
                    cleaned = cleaned.lstrip("+")
                    if not cleaned:
                        continue
                    try:
                        val = float(cleaned)
                        nums.append(val)
                    except ValueError:
                        continue

                if nums:
                    total_sum += nums[0]
                    gst_sum += nums[-1]

            results["TOTAL"] = round(total_sum, 2) if total_sum else None
            results["TOTAL GST"] = round(gst_sum, 2) if gst_sum else None

    except Exception as e:
        st.error(f"❌ Error processing {filename}: {e}")

    return results


# --------- Streamlit App ---------
st.title("📄 PDF Extractor")

uploaded_files = st.file_uploader(
    "Upload PDF files", type="pdf", accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    progress = st.progress(0)  # progress bar
    status = st.empty()        # placeholder for filename
    
    for i, uploaded_file in enumerate(uploaded_files, start=1):
        filename = uploaded_file.name
        status.text(f"🔄 Processing: {filename}")
        with st.spinner(f"Extracting data from {filename}..."):
            results = extract_from_pdf(uploaded_file, filename)
            all_data.append(results)

        progress.progress(i / len(uploaded_files))

    # Convert to DataFrame
    df = pd.DataFrame(all_data)

    # Show extracted data
    st.subheader("✅ Extracted Data")
    st.dataframe(df)

    # Save to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:   
        df.to_excel(writer, index=False, sheet_name="Invoices")
    output.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=output,
        file_name="invoices_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
