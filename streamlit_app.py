# streamlit run streamlit_app.py

import streamlit as st
from io import BytesIO

# Check if python-docx is available
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(page_title="Word Template Filler", layout="centered")
st.title("📄 Word Template Auto-Filler")

# Show error if docx module is not available
if not DOCX_AVAILABLE:
    st.error("""
    🚨 **Missing Required Package**
    
    The `python-docx` package is not installed. Please install it by running:
    ```
    pip install python-docx
    ```
    
    Or if deploying to Streamlit Cloud, create a `requirements.txt` file with:
    ```
    streamlit
    python-docx
    ```
    """)
    st.stop()

# Upload Word template
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if template_file:
    with st.form("input_form"):
        st.subheader("🔹 General Info")
        client_name = st.text_input("Client Name")
        contact_person = st.text_input("Contact Person")
        offer_number = st.text_input("Offer Number")
        quotation_date = st.date_input("Quotation Date")
        mobilization_days = st.text_input("Mobilization Days")
        inspection_days = st.text_input("Inspection Days")
        total_days = st.text_input("Total Duration (Days)")
        pipe_diameter = st.text_input("Pipe Diameter (e.g., 16inch)")

        st.subheader("🔹 Segment Lengths")
        segment_lengths = [st.text_input(f"Segment {i+1} Length (e.g., 98.08 KM)") for i in range(7)]

        st.subheader("🔹 Segment Prices")
        prices = [st.text_input(f"Price for Segment {i+1}") for i in range(7)]

        st.subheader("🔹 Additional Costs")
        mfl_tool_rerun = st.text_input("MFL Tool Re-run Cost")
        egp_tool_rerun = st.text_input("EGP Tool Re-run Cost")
        egp_additional_mob = st.text_input("Additional EGP Mobilization Cost")
        mfl_additional_mob = st.text_input("Additional MFL/TFI Mobilization Cost")
        personnel_additional_mob = st.text_input("Additional Mobilization Personnel Cost")
        mfl_standby_rate = st.text_input("MFL/TFI Standby Rate")
        egp_standby_rate = st.text_input("EGP Standby Rate")
        personnel_standby_rate = st.text_input("Personnel Standby Rate")

        submitted = st.form_submit_button("Generate Document")

    if submitted:
        try:
            # Load template
            doc = Document(template_file)
            replacements = {
                "<<client_name>>": client_name or "",
                "<<contact_person>>": contact_person or "",
                "<<offer_number>>": offer_number or "",
                "<<quotation_date>>": quotation_date.strftime("%d/%m/%Y") if quotation_date else "",
                "<<mobilization_days>>": mobilization_days or "",
                "<<inspection_days>>": inspection_days or "",
                "<<total_days>>": total_days or "",
                "<<pipe_diameter>>": pipe_diameter or "",
                "<<mfl_tool_rerun>>": mfl_tool_rerun or "",
                "<<egp_tool_rerun>>": egp_tool_rerun or "",
                "<<egp_additional_mob>>": egp_additional_mob or "",
                "<<mfl_additional_mob>>": mfl_additional_mob or "",
                "<<personnel_additional_mob>>": personnel_additional_mob or "",
                "<<mfl_standby_rate>>": mfl_standby_rate or "",
                "<<egp_standby_rate>>": egp_standby_rate or "",
                "<<personnel_standby_rate>>": personnel_standby_rate or "",
            }

            for i in range(7):
                segment = segment_lengths[i] if i < len(segment_lengths) else ""
                price = prices[i] if i < len(prices) else ""
                replacements[f"<<segment_{i+1}>>"] = f"{pipe_diameter}x{segment}" if pipe_diameter and segment else segment
                replacements[f"<<price_segment_{i+1}>>"] = price

            # Replace placeholders in paragraphs
            for para in doc.paragraphs:
                for key, val in replacements.items():
                    if key in para.text:
                        para.text = para.text.replace(key, val)

            # Replace placeholders in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, val in replacements.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, val)

            # Save to buffer and provide download
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Document generated successfully!")
            st.download_button(
                label="📥 Download Filled Document",
                data=buffer,
                file_name=f"Filled_Document_{offer_number or 'template'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"❌ Error processing document: {str(e)}")
            st.info("Make sure your Word template file is valid and not corrupted.")
else:
    st.info("👆 Please upload a Word template (.docx) file to get started.")
