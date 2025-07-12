import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="📄 BPCL Quarterly PDF Extractor")

st.title("📄 BPCL Quarterly PDF Data Extractor")

st.write("""
Upload BPCL quarterly PDF report(s).  
This app will extract key performance indicators using **tables and plain text**, then generate a combined Excel file.
""")

# Upload BPCL PDFs
bpcl_pdfs = st.file_uploader("Upload BPCL PDF(s)", type=["pdf"], accept_multiple_files=True)

process = st.button("🚀 Process BPCL PDFs")

if process and bpcl_pdfs:

    def parse_bpcl_pdf(pdf_file):
        all_text = ""
        table_data = {}

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"

                # Try to find tables
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # SAFELY join: skip None
                        joined = " ".join([cell for cell in row if cell]).lower()

                        if "crude throughput" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["Crude Throughput (MMT)"] = m.group(1)
                        if "distillate yield" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["Distillate Yield (%)"] = m.group(1)
                        if "sko" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["SKO (MMT)"] = m.group(1)
                        if "atf" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["ATF (MMT)"] = m.group(1)
                        if "others" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["Others (MMT)"] = m.group(1)
                        if "exports" in joined:
                            for cell in row:
                                if cell:
                                    m = re.search(r"([\d\.]+)", cell)
                                    if m:
                                        table_data["Exports (MMT)"] = m.group(1)

        # Extract from plain text for fields not in tables
        text_data = {
            "Sl.No": None,
            "Company": "BPCL",
            "Fiscal Year": re.search(r"Fiscal Year\s*:\s*(\d{4}-\d{4})", all_text),
            "Quarter": re.search(r"Quarter\s*:\s*(Q[1-4])", all_text),
            "Duration": re.search(r"Duration\s*:\s*(.*?)\n", all_text),
            "Date": re.search(r"Date\s*:\s*([\d-]+)", all_text),
            "MR Throughput (MMT)": re.search(r"MR Throughput.*?([\d\.]+)", all_text),
            "KR Throughput (MMT)": re.search(r"KR Throughput.*?([\d\.]+)", all_text),
            "BR Throughput (MMT)": re.search(r"BR Throughput.*?([\d\.]+)", all_text),
            "HS crude (%)": re.search(r"HS crude.*?([\d\.]+)", all_text),
            "Utilisation (%)": re.search(r"Utilisation.*?([\d\.]+)", all_text),
            "Domestic Sales (MMT)": re.search(r"Domestic Sales.*?([\d\.]+)", all_text),
            "LPG Sales (MMT)": re.search(r"LPG Sales.*?([\d\.]+)", all_text),
            "MS Sales (MMT)": re.search(r"MS Sales.*?([\d\.]+)", all_text),
            "HSD Sales (MMT)": re.search(r"HSD Sales.*?([\d\.]+)", all_text),
            "Pipeline Throughput (MMT)": re.search(r"Pipeline Throughput.*?([\d\.]+)", all_text),
            "Gross Margins ($/bbl)": re.search(r"Gross Margins[^$]*\$?.*?([\d\.]+)", all_text),
            "Gross Margins - MR ($/bbl)": re.search(r"Gross Margins - MR[^$]*\$?.*?([\d\.]+)", all_text),
            "Gross Margins - KR ($/bbl)": re.search(r"Gross Margins - KR[^$]*\$?.*?([\d\.]+)", all_text),
            "Gross Margins - BR ($/bbl)": re.search(r"Gross Margins - BR[^$]*\$?.*?([\d\.]+)", all_text),
            "Revenue from operations (₹ Crores)": re.search(r"Revenue from operations.*?₹\s*([\d,]+)", all_text),
            "Cost of materials consumed (₹ Crores)": re.search(r"Cost of materials.*?₹\s*([\d,]+)", all_text),
            "Purchase of stock-in-trade (₹ Crores)": re.search(r"Purchase of stock.*?₹\s*([\d,]+)", all_text),
            "Change in inventories (₹ Crores)": re.search(r"Change in inventories.*?₹\s*([\d,]+)", all_text),
            "PBT (₹ Crores)": re.search(r"PBT.*?₹\s*([\d,]+)", all_text),
            "Domestic market share of POL": re.search(r"Domestic market share.*?([\d\.]+)", all_text),
            "Retail Outlets": re.search(r"Retail Outlets.*?([\d,]+)", all_text),
            "LPG Distributionship": re.search(r"LPG Distributorship.*?([\d,]+)", all_text),
            "CNG facilities at ROs": re.search(r"CNG facilities.*?([\d,]+)", all_text),
            "Aviation Service stations": re.search(r"Aviation Service.*?([\d,]+)", all_text),
            "LPG Consumers (million)": re.search(r"LPG Consumers.*?([\d\.]+)", all_text)
        }

        # Combine: prefer table data if present
        combined = {}
        for key in text_data.keys():
            if key in table_data and table_data[key] is not None:
                combined[key] = table_data[key]
            else:
                v = text_data[key]
                if isinstance(v, str):
                    combined[key] = v
                elif v:
                    combined[key] = v.group(1)
                else:
                    combined[key] = None

        return combined


    bpcl_rows = []
    for f in bpcl_pdfs:
        bpcl_rows.append(parse_bpcl_pdf(f))

    bpcl_df = pd.DataFrame(bpcl_rows)
    bpcl_df["Sl.No"] = range(1, len(bpcl_df) + 1)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        bpcl_df.to_excel(writer, sheet_name="BPCL", index=False)

    st.success("✅ Done! Download your BPCL Excel below.")
    st.download_button(
        "📥 Download BPCL Excel",
        data=output.getvalue(),
        file_name="BPCL_Quarterly_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    if process:
        st.warning("Please upload at least one BPCL PDF.")

