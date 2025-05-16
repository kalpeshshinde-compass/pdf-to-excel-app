import pandas as pd
import fitz  # PyMuPDF
import streamlit as st
from io import BytesIO
from datetime import datetime

st.title("PDF to Excel Expense Extractor")

uploaded_pdf = st.file_uploader("Upload Expense Report PDF", type="pdf")

if uploaded_pdf:
    doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
    text = "\n".join([page.get_text() for page in doc])

    # Extract fixed fields
    def extract_field(field_name, text):
        for line in text.splitlines():
            if field_name in line:
                return line.split(":")[-1].strip()
        return ""

    employee_name = extract_field("Employee Name", text)
    employee_id = extract_field("Employee ID", text)
    report_name = extract_field("Report Name", text)
    report_date = extract_field("Report Date", text)

    # Extract table data
    lines = text.splitlines()
    data = []
    for i, line in enumerate(lines):
        if "Uber" in line and "Out of" in line:
            try:
                transaction_date = lines[i - 1].strip()
                vendor = "Uber"
                net_amount = lines[i + 1].split("$")[1].strip()
                tax_amount = lines[i + 2].split("$")[1].strip()
                total_amount = float(net_amount) + float(tax_amount)

                data.append({
                    "Employee Name": employee_name,
                    "Employee ID": employee_id,
                    "Report Name": report_name,
                    "Report Date": report_date,
                    "Vendor": vendor,
                    "Transaction Date": transaction_date,
                    "Net Adjusted Reclaim Amount": net_amount,
                    "Tax Reclaim Amount": tax_amount,
                    "Total": f"{total_amount:.2f}"
                })
            except Exception as e:
                st.error(f"Error processing line {i}: {e}")

    if data:
        df = pd.DataFrame(data)
        st.dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name=f"expense_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No valid entries found in the PDF.")
