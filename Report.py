import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="📄 BPCL Quarterly PDF Extractor")

st.title("📄 BPCL Quarterly PDF Data Extractor")

st.write("""
Upload BPCL quarterly PDF report(s).  
This app extracts values from tables with multiple columns (Q4 FY18, Q4 FY17, FY18, FY17).  
By default, we pick **first column after the label**.
""")

# Let user choose which column
col_option = st.selectbox(
    "📊 Which column to pick?",
    ["Q4 FY18 (1st number)", "Q4 FY17 (2nd number)", "FY18 (3rd number)", "FY17 (4th number)"]
)

col_map = {
    "Q4 FY18 (1st number)": 1,
    "Q4 FY17 (2nd number)": 2,
    "FY18 (3rd number)": 3,
    "FY17 (4th number)": 4
}
col_idx = col_map[col_option]

# Upload BPCL PDFs
bpcl_pdfs = st.file_uploader("Upload BPCL PDF(s)", type=["pdf"], accept_multiple_files=True)

process = st.button("🚀 Process BPCL PDFs")

if process and bpcl_pdfs:

    def parse_bpcl_pdf(pdf_file, col_idx):
        result = {"Company": "BPCL"}

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row is None: continue
                        row = [cell for cell in row if cell]

                        joined = " ".join(row).lower()

                        def get_value_if_match(prefixes):
                            for p in prefixes:
                                if joined.startswith(p.lower()):
                                    if len(row) > col_idx:
                                        return row[col_idx]
                            return None

                        if val := get_value_if_match(["Refinery Crude Throughput MMT"]):
                            result["Crude Throughput (MMT)"] = val
                        if val := get_value_if_match(["- MR MMT", "MR MMT"]):
                            result["MR Throughput (MMT)"] = val
                        if val := get_value_if_match(["- KR MMT", "KR MMT"]):
                            result["KR Throughput (MMT)"] = val
                        if val := get_value_if_match(["Distillate Yield %"]):
                            result["Distillate Yield (%)"] = val
                        if val := get_value_if_match(["High Sulphur"]):
                            result["HS crude (%)"] = val
                        if val := get_value_if_match(["- LPG MMT"]):
                            result["LPG Sales (MMT)"] = val
                        if val := get_value_if_match(["- MS MMT"]):
                            result["MS Sales (MMT)"] = val
                        if val := get_value_if_match(["- HSD MMT"]):
                            result["HSD Sales (MMT)"] = val
                        if val := get_value_if_match(["- SKO MMT"]):
                            result["SKO (MMT)"] = val
                        if val := get_value_if_match(["- ATF MMT"]):
                            result["ATF (MMT)"] = val
                        if val := get_value_if_match(["- Others MMT"]):
                            result["Others (MMT)"] = val
                        if val := get_value_if_match(["b. Exports MMT"]):
                            result["Exports (MMT)"] = val
                        if val := get_value_if_match(["Total Domestic MMT"]):
                            result["Domestic Sales (MMT)"] = val
                        if val := get_value_if_match(["Total Sales MMT"]):
                            result["Total Sales (MMT)"] = val
                        if val := get_value_if_match(["GRM (BPCL) US"]):
                            result["Gross Margins ($/bbl)"] = val
                        if val := get_value_if_match(["GRM (Mumbai Refinery) US"]):
                            result["Gross Margins - MR ($/bbl)"] = val
                        if val := get_value_if_match(["GRM (Kochi Refinery) US"]):
                            result["Gross Margins - KR ($/bbl)"] = val

        return result


    bpcl_rows = []
    for f in bpcl_pdfs:
        data = parse_bpcl_pdf(f, col_idx)
        bpcl_rows.append(data)

    bpcl_df = pd.DataFrame(bpcl_rows)
    bpcl_df.insert(0, "Sl.No", range(1, len(bpcl_df)+1))

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
        st.warning("Upload at least one BPCL PDF.")
