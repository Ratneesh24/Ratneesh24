import streamlit as st
from datetime import date
import tempfile
import os

from mill_plan_generator import generate_daily_plan

st.set_page_config(page_title="Mill Plan Generator", layout="wide")

st.title("🏭 Mill Plan Generator")
st.write("Upload the WIP Excel file and generate the planning workbook.")

uploaded_file = st.file_uploader(
    "Upload Narrow Data Coil Stage File",
    type=["xlsx"]
)

plan_date = st.date_input("Plan Date", value=date.today())
days = st.number_input("Number of Days", min_value=1, max_value=30, value=1)

if uploaded_file is not None:
    st.success("File uploaded successfully.")

    if st.button("Generate Mill Plan"):
        with st.spinner("Generating plan..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, uploaded_file.name)
                output_path = os.path.join(tmpdir, "mill_plan_output.xlsx")

                # Save uploaded file
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Run your existing logic
                generate_daily_plan(
                    wip_file=input_path,
                    plan_date=plan_date,
                    output_file=output_path,
                    days=days,
                )

                # Read output file
                with open(output_path, "rb") as f:
                    output_bytes = f.read()

        st.success("Mill plan generated successfully!")

        st.download_button(
            label="📥 Download Mill Plan",
            data=output_bytes,
            file_name=f"Mill_Plan_{plan_date.strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
