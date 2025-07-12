import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="OMC Quarterly Data Extractor")

st.title("📊 Quarterly PDF Data Extractor for OMCs (HPCL, BPCL, IOCL, RIL)")

st.write("""
Upload quarterly PDF reports for each company below.  
The app will extract the required KPIs, merge all data, and generate a combined Excel file.
""")

# Uploaders for each company
hpcl_pdfs = st.file_uploader("📄 Upload HPCL PDF(s)", type=["pdf"], accept_multiple_files=True)
bpcl_pdfs = st.file_uploader("📄 Upload BPCL PDF(s)", type=["pdf"], accept_multiple_files=True)
iocl_pdfs = st.file_uploader("📄 Upload IOCL PDF(s)", type=["pdf"], accept_multiple_files=True)
ril_pdfs = st.file_uploader("📄 Upload RIL PDF(s)", type=["pdf"], accept_multiple_files=True)

process = st.button("🚀 Process PDFs & Generate Excel")

if process:
    def parse_pdf(pdf_file, company):
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        if company == "HPCL":
            data = {
                "Sl.No": None,
                "Company": "HPCL",
                "Fiscal Year": re.search(r"Fiscal Year\s*:\s*(\d{4}-\d{4})", text),
                "Quarter": re.search(r"Quarter\s*:\s*(Q[1-4])", text),
                "Duration": re.search(r"Duration\s*:\s*(.*?)\n", text),
                "Date": re.search(r"Date\s*:\s*([\d-]+)", text),
                "Crude Throughput (MMT)": re.search(r"Crude Throughput.*?([\d\.]+)", text),
                "Utilisation (%)": re.search(r"Utilisation.*?([\d\.]+)", text),
                "Domestic Sales (MMT)": re.search(r"Domestic Sales.*?([\d\.]+)", text),
                "Exports (MMT)": re.search(r"Exports.*?([\d\.]+)", text),
                "Pipeline Throughput (MMT)": re.search(r"Pipeline Throughput.*?([\d\.]+)", text),
                "Gross Margins ($/bbl)": re.search(r"Gross Margins.*?([\d\.]+)", text),
                "Sale of products.*?₹\s*([\d,]+)": re.search(r"Sale of products.*?₹\s*([\d,]+)", text),
                "Cost of material consumed.*?₹\s*([\d,]+)": re.search(r"Cost of material.*?₹\s*([\d,]+)", text),
                "Purchases of stock-in-trade.*?₹\s*([\d,]+)": re.search(r"Purchases of stock.*?₹\s*([\d,]+)", text),
                "Change in inventories.*?₹\s*([\d,]+)": re.search(r"Change in inventories.*?₹\s*([\d,]+)", text),
                "PBT.*?₹\s*([\d,]+)": re.search(r"PBT.*?₹\s*([\d,]+)", text),
                "Domestic market share of POL": re.search(r"Domestic market share.*?([\d\.]+)", text),
                "Retail Outlets": re.search(r"Retail Outlets.*?([\d,]+)", text),
                "LPG Distributionship": re.search(r"LPG Distributorship.*?([\d,]+)", text),
                "SKO/LDO Dealership": re.search(r"SKO.*?([\d,]+)", text),
                "Lube Distributors.*?": re.search(r"Lube Distributors.*?([\d,]+)", text),
                "Mobile Dispensers": re.search(r"Mobile Dispensers.*?([\d,]+)", text),
                "CNG facilities at ROs": re.search(r"CNG.*?([\d,]+)", text),
                "EV Charging facilities at Ros": re.search(r"EV Charging.*?([\d,]+)", text),
                "LPG Consumers.*?million": re.search(r"LPG Consumers.*?([\d\.]+)", text)
            }

        elif company == "BPCL":
            data = {
                "Sl.No": None,
                "Company": "BPCL",
                "Fiscal Year": re.search(r"Fiscal Year\s*:\s*(\d{4}-\d{4})", text),
                "Quarter": re.search(r"Quarter\s*:\s*(Q[1-4])", text),
                "Duration": re.search(r"Duration\s*:\s*(.*?)\n", text),
                "Date": re.search(r"Date\s*:\s*([\d-]+)", text),
                "Crude Throughput (MMT)": re.search(r"Crude Throughput.*?([\d\.]+)", text),
                "MR Throughput (MMT)": re.search(r"MR Throughput.*?([\d\.]+)", text),
                "KR Throughput (MMT)": re.search(r"KR Throughput.*?([\d\.]+)", text),
                "BR Throughput (MMT)": re.search(r"BR Throughput.*?([\d\.]+)", text),
                "Distillate Yield (%)": re.search(r"Distillate Yield.*?([\d\.]+)", text),
                "HS crude (%)": re.search(r"HS crude.*?([\d\.]+)", text),
                "Utilisation (%)": re.search(r"Utilisation.*?([\d\.]+)", text),
                "Domestic Sales (MMT)": re.search(r"Domestic Sales.*?([\d\.]+)", text),
                "LPG Sales (MMT)": re.search(r"LPG Sales.*?([\d\.]+)", text),
                "MS Sales (MMT)": re.search(r"MS Sales.*?([\d\.]+)", text),
                "HSD Sales (MMT)": re.search(r"HSD Sales.*?([\d\.]+)", text),
                "SKO (MMT)": re.search(r"SKO.*?([\d\.]+)", text),
                "ATF (MMT)": re.search(r"ATF.*?([\d\.]+)", text),
                "Others (MMT)": re.search(r"Others.*?([\d\.]+)", text),
                "Exports (MMT)": re.search(r"Exports.*?([\d\.]+)", text),
                "Pipeline Throughput (MMT)": re.search(r"Pipeline.*?([\d\.]+)", text),
                "Gross Margins ($/bbl)": re.search(r"Gross Margins.*?([\d\.]+)", text),
                "Gross Margins - MR ($/bbl)": re.search(r"Gross Margins - MR.*?([\d\.]+)", text),
                "Gross Margins - KR ($/bbl)": re.search(r"Gross Margins - KR.*?([\d\.]+)", text),
                "Gross Margins - BR ($/bbl)": re.search(r"Gross Margins - BR.*?([\d\.]+)", text),
                "Revenue from operations.*?₹\s*([\d,]+)": re.search(r"Revenue from operations.*?₹\s*([\d,]+)", text),
                "Cost of materials.*?₹\s*([\d,]+)": re.search(r"Cost of materials.*?₹\s*([\d,]+)", text),
                "Purchase of stock.*?₹\s*([\d,]+)": re.search(r"Purchase of stock.*?₹\s*([\d,]+)", text),
                "Change in inventories.*?₹\s*([\d,]+)": re.search(r"Change in inventories.*?₹\s*([\d,]+)", text),
                "PBT.*?₹\s*([\d,]+)": re.search(r"PBT.*?₹\s*([\d,]+)", text),
                "Domestic market share of POL": re.search(r"Domestic market share.*?([\d\.]+)", text),
                "Retail Outlets": re.search(r"Retail Outlets.*?([\d,]+)", text),
                "LPG Distributionship": re.search(r"LPG Distributorship.*?([\d,]+)", text),
                "CNG facilities at ROs": re.search(r"CNG facilities.*?([\d,]+)", text),
                "Aviation Service stations": re.search(r"Aviation Service.*?([\d,]+)", text),
                "LPG Consumers.*?million": re.search(r"LPG Consumers.*?([\d\.]+)", text)
            }

        elif company == "IOCL":
            data = {
                "Sl.No": None,
                "Company": "IOCL",
                "Fiscal Year": re.search(r"Fiscal Year\s*:\s*(\d{4}-\d{4})", text),
                "Quarter": re.search(r"Quarter\s*:\s*(Q[1-4])", text),
                "Duration": re.search(r"Duration\s*:\s*(.*?)\n", text),
                "Date": re.search(r"Date\s*:\s*([\d-]+)", text),
                "Crude Throughput (MMT)": re.search(r"Crude Throughput.*?([\d\.]+)", text),
                "Utilisation (%)": re.search(r"Utilisation.*?([\d\.]+)", text),
                "Distillate yield (%)": re.search(r"Distillate yield.*?([\d\.]+)", text),
                "F&L (%)": re.search(r"F&L.*?([\d\.]+)", text),
                "Domestic Sales (MMT)": re.search(r"Domestic Sales.*?([\d\.]+)", text),
                "Exports (MMT)": re.search(r"Exports.*?([\d\.]+)", text),
                "Pipeline Throughput (MMT)": re.search(r"Pipeline.*?([\d\.]+)", text),
                "Gross Margins ($/bbl)": re.search(r"Gross Margins.*?([\d\.]+)", text),
                "Revenue from operations.*?₹\s*([\d,]+)": re.search(r"Revenue from operations.*?₹\s*([\d,]+)", text),
                "Cost of material.*?₹\s*([\d,]+)": re.search(r"Cost of material.*?₹\s*([\d,]+)", text),
                "Purchases of stock.*?₹\s*([\d,]+)": re.search(r"Purchases of stock.*?₹\s*([\d,]+)", text),
                "Change in inventories.*?₹\s*([\d,]+)": re.search(r"Change in inventories.*?₹\s*([\d,]+)", text),
                "PBT.*?₹\s*([\d,]+)": re.search(r"PBT.*?₹\s*([\d,]+)", text),
                "Domestic market share of POL": re.search(r"Domestic market share.*?([\d\.]+)", text),
                "Retail Outlets": re.search(r"Retail Outlets.*?([\d,]+)", text),
                "LPG Distributionship": re.search(r"LPG Distributorship.*?([\d,]+)", text),
                "SKO/LDO Dealership": re.search(r"SKO/LDO.*?([\d,]+)", text),
                "Lube Distributors.*?": re.search(r"Lube Distributors.*?([\d,]+)", text),
                "Mobile Dispensers": re.search(r"Mobile Dispensers.*?([\d,]+)", text),
                "CNG facilities at ROs": re.search(r"CNG facilities.*?([\d,]+)", text),
                "Aviation Fuel stations": re.search(r"Aviation Fuel.*?([\d,]+)", text),
                "LPG Consumers.*?million": re.search(r"LPG Consumers.*?([\d\.]+)", text),
                "Russian crude %": re.search(r"Russian crude.*?([\d\.]+)", text)
            }

        elif company == "RIL":
            data = {
                "Sl.No": None,
                "Company": "RIL",
                "Fiscal Year": re.search(r"Fiscal Year\s*:\s*(\d{4}-\d{4})", text),
                "Quarter": re.search(r"Quarter\s*:\s*(Q[1-4])", text),
                "Duration": re.search(r"Duration\s*:\s*(.*?)\n", text),
                "Date": re.search(r"Date\s*:\s*([\d-]+)", text),
                "Crude Throughput (MMT)": re.search(r"Crude Throughput.*?([\d\.]+)", text),
                "Utilisation (%)": re.search(r"Utilisation.*?([\d\.]+)", text),
                "Total Sales (MMT)": re.search(r"Total Sales.*?([\d\.]+)", text),
                "Gross Margins ($/bbl)": re.search(r"Gross Margins.*?([\d\.]+)", text),
                "Revenue from operations.*?₹\s*([\d,]+)": re.search(r"Revenue from operations.*?₹\s*([\d,]+)", text),
                "Exports.*?₹\s*([\d,]+)": re.search(r"Exports.*?₹\s*([\d,]+)", text),
                "EBIDTA.*?₹\s*([\d,]+)": re.search(r"EBIDTA.*?₹\s*([\d,]+)", text),
                "EBIDTA %": re.search(r"EBIDTA %.*?([\d\.]+)", text),
                "Retail outlets": re.search(r"Retail outlets.*?([\d,]+)", text),
                "HSD Sales Y-o-Y (%)": re.search(r"HSD Sales.*?([\d\.]+)", text),
                "MS Sales Y-o-Y (%)": re.search(r"MS Sales.*?([\d\.]+)", text),
                "ATF sales.*?([\d\.]+)": re.search(r"ATF sales.*?([\d\.]+)", text),
                "Charging points": re.search(r"Charging points.*?([\d,]+)", text),
                "Unique sites": re.search(r"Unique sites.*?([\d,]+)", text),
                "CBG Network": re.search(r"CBG Network.*?([\d,]+)", text)
            }

        clean_data = {}
        for k, v in data.items():
            clean_data[k] = v.group(1) if v else None

        return clean_data

    # Store results
    hpcl_data, bpcl_data, iocl_data, ril_data = [], [], [], []

    for f in hpcl_pdfs:
        hpcl_data.append(parse_pdf(f, "HPCL"))
    for f in bpcl_pdfs:
        bpcl_data.append(parse_pdf(f, "BPCL"))
    for f in iocl_pdfs:
        iocl_data.append(parse_pdf(f, "IOCL"))
    for f in ril_pdfs:
        ril_data.append(parse_pdf(f, "RIL"))

    # Convert to DF
    hpcl_df = pd.DataFrame(hpcl_data)
    bpcl_df = pd.DataFrame(bpcl_data)
    iocl_df = pd.DataFrame(iocl_data)
    ril_df = pd.DataFrame(ril_data)

    # Add Serial Numbers
    hpcl_df["Sl.No"] = range(1, len(hpcl_df) + 1)
    bpcl_df["Sl.No"] = range(1, len(bpcl_df) + 1)
    iocl_df["Sl.No"] = range(1, len(iocl_df) + 1)
    ril_df["Sl.No"] = range(1, len(ril_df) + 1)

    # Output to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        hpcl_df.to_excel(writer, sheet_name="HPCL", index=False)
        bpcl_df.to_excel(writer, sheet_name="BPCL", index=False)
        iocl_df.to_excel(writer, sheet_name="IOCL", index=False)
        ril_df.to_excel(writer, sheet_name="RIL", index=False)

    st.success("✅ Data extraction done! Download your Excel below.")
    st.download_button(
        "📥 Download Combined Excel",
        data=output.getvalue(),
        file_name="OMCs_Quarterly_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
