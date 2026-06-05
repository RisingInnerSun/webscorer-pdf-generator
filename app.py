from pathlib import Path
import tempfile

import streamlit as st

from race_results_to_pdf import create_pdf


st.set_page_config(page_title="Webscorer PDF Generator", layout="centered")

st.title("Webscorer PDF Generator")
st.write("Upload a Webscorer Excel results file, optionally enter a race name/date, then generate a PDF.")

uploaded_file = st.file_uploader("Select Webscorer Excel file", type=["xlsx", "xls"])

race_name = st.text_input(
    "Race name and results title (optional)",
    help="Leave blank to use the Excel filename as the race name.",
)

race_date = st.text_input(
    "Race date (optional)",
    help="Leave blank to detect the date from the workbook, or use the file modified date as a fallback.",
)

output_filename = st.text_input("Output PDF filename", value="webscorer_results.pdf")
if not output_filename.lower().endswith(".pdf"):
    output_filename = f"{output_filename}.pdf"

if uploaded_file is not None:
    st.info(f"Selected file: {uploaded_file.name}")

    if st.button("Generate PDF", type="primary"):
        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                tmp_dir_path = Path(tmp_dir)
                excel_path = tmp_dir_path / uploaded_file.name
                pdf_path = tmp_dir_path / output_filename

                excel_path.write_bytes(uploaded_file.getbuffer())

                create_pdf(
                    excel_file=str(excel_path),
                    output_pdf=str(pdf_path),
                    race_name=race_name.strip() or None,
                    race_date=race_date.strip() or None,
                )

                pdf_bytes = pdf_path.read_bytes()

            st.success("PDF created successfully.")
            st.download_button(
                label="Download PDF",
                data=pdf_bytes,
                file_name=output_filename,
                mime="application/pdf",
            )

        except Exception as exc:
            st.error("PDF generation failed.")
            st.exception(exc)
else:
    st.warning("Upload an Excel file to begin.")
