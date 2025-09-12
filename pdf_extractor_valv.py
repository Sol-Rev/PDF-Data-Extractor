import streamlit as st
import pandas as pd
import re
import pdfplumber
from pathlib import Path
from PyPDF2 import PdfReader
import io
from openpyxl import load_workbook

# -------- Function 1: Extract invoice summary --------
def data_rest(file, filename):
    try:
        reader = PdfReader(file)
    except Exception as e:
        st.error(f"❌ Error opening PDF {filename}: {e}")
        return None

    fields = ['Invoice Number', 'Store Invoice Nbr', 'Svc Dt:', 'AUTHNUM', 'VIN']
    totals_labels = ['PRE-TAX AMOUNT', 'TAX', 'INVOICE TOTAL', 'DISCOUNTABLE TOTAL', 'FLEET DISCOUNT', 'TOTAL DUE']
    all_invoices = []

    for page in reader.pages:
        text = page.extract_text().split('\n')
        invoice_data = {}
        index_map = [i+1 for i, x in enumerate(text) if any(field in x for field in fields)]

        i = 0
        while i < len(text):
            line = text[i].strip()

            if 'Invoice Number' in line:
                if invoice_data:
                    invoice_data['Source File'] = filename
                    invoice_data['PO.x'] = invoice_data.get('AUTHNUM', '')
                    all_invoices.append(invoice_data)
                    invoice_data = {}

                if i+1 < len(text):
                    next_line = text[i+1].strip()
                    match = re.search(r"\d+", next_line)
                    invoice_data['Invoice Number'] = match.group() if match else next_line
                i += 2
                continue

            if i in index_map:
                line_val = text[i].strip()

                date = re.search(r"\d{4}-\d{2}-\d{2}", line_val)
                if date:
                    invoice_data['Svc Dt:'] = date.group()
                    i += 1
                    continue

                vin = re.search(r"[A-HJ-NPR-Z0-9]{17}", line_val)
                if vin:
                    invoice_data['VIN'] = vin.group()
                    i += 1
                    continue

                match = re.search(r"\d+", line_val)
                if match:
                    for f in ['Store Invoice Nbr', 'AUTHNUM']:
                        if f not in invoice_data:
                            invoice_data[f] = match.group()
                            break
                    i += 1
                    continue

                i += 1
                continue

            if 'PRE-TAX AMOUNT' in line:
                relevant_lines = text[i+1:]
                numbers = []
                for l in relevant_lines:
                    l = l.strip()
                    if re.match(r'^-?\d+(?:\.\d+)?$', l):
                        numbers.append(float(l))
                    if len(numbers) == len(totals_labels):
                        break
                if len(numbers) == len(totals_labels):
                    for k, v in zip(totals_labels, numbers):
                        invoice_data[k] = v
                i += len(numbers) + 1
                continue

            i += 1

        if invoice_data:
            invoice_data['Source File'] = filename
            invoice_data['PO.x'] = invoice_data.get('AUTHNUM', '')
            all_invoices.append(invoice_data)

    if all_invoices:
        return pd.DataFrame(all_invoices, columns=['Source File'] + fields + totals_labels + ['PO.x'])
    else:
        return None


# -------- Function 2: Extract invoice table --------
def data_table(file, filename):
    data_rows = []

    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text().split('\n')

            try:
                header_index = [i for i, x in enumerate(text) if 'Invoice Store Invoice' in x][0]
            except IndexError:
                continue

            for line in text[header_index + 2:]:
                if not re.search(r"\d{4}-\d{2}-\d{2}", line):
                    break

                clean_line = re.sub(r"_+", "", line).strip()
                parts = clean_line.split()
                if len(parts) >= 4:
                    invoice, store_invoice, service_date, amount = parts[:4]
                    data_rows.append([invoice, store_invoice, service_date, amount, filename])

    if data_rows:
        return pd.DataFrame(data_rows, columns=["Invoice", "Store Invoice", "Service Date", "Amount", "Source File"])
    else:
        return None


# -------- Save both to Excel --------
def save_to_excel(dfs_rest, dfs_table):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if dfs_rest:
            pd.concat(dfs_rest, ignore_index=True).to_excel(writer, sheet_name="Invoice_Summary", index=False)
        if dfs_table:
            pd.concat(dfs_table, ignore_index=True).to_excel(writer, sheet_name="Invoice_Table", index=False)

    output.seek(0)
    return output


# -------- Streamlit App --------
st.title("📑 PDF Invoice Extractor")

uploaded_files = st.file_uploader("Upload one or more PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    dfs_rest, dfs_table = [], []

    with st.spinner("⏳ Extracting data from uploaded PDFs..."):
        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            st.info(f"Processing {filename}...")  # optional, shows current file

            # Extract invoice summary
            df_rest = data_rest(uploaded_file, filename)
            if df_rest is not None:
                dfs_rest.append(df_rest)

            uploaded_file.seek(0)  # reset pointer

            # Extract invoice table
            df_table = data_table(uploaded_file, filename)
            if df_table is not None:
                dfs_table.append(df_table)

    st.success("✅ Extraction complete for all uploaded files!")

    # Preview extracted data
    if dfs_rest:
        st.subheader("📑 Extracted Invoice Summary")
        df_rest_all = pd.concat(dfs_rest, ignore_index=True)
        st.write(f"Total rows: {len(df_rest_all)}")
        st.dataframe(df_rest_all)

    if dfs_table:
        st.subheader("📊 Extracted Invoice Table")
        df_table_all = pd.concat(dfs_table, ignore_index=True)
        st.write(f"Total rows: {len(df_table_all)}")
        st.dataframe(df_table_all)

    
    if dfs_rest or dfs_table:
        excel_file = save_to_excel(dfs_rest, dfs_table)
        st.success("✅ Extraction complete!")
        st.download_button(
            label="📥 Download Excel File",
            data=excel_file,
            file_name="invoices_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No data extracted from the uploaded PDFs.")
