# COPY THIS
# streamlit run streamlit_app.py



import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Word Template Filler", layout="centered")
st.title("📄 Word Template Auto-Filler")

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
        # Load template
        doc = Document(template_file)
        replacements = {
            "<<client_name>>": client_name,
            "<<contact_person>>": contact_person,
            "<<offer_number>>": offer_number,
            "<<quotation_date>>": quotation_date.strftime("%d/%m/%Y"),
            "<<mobilization_days>>": mobilization_days,
            "<<inspection_days>>": inspection_days,
            "<<total_days>>": total_days,
            "<<pipe_diameter>>": pipe_diameter,
            "<<mfl_tool_rerun>>": mfl_tool_rerun,
            "<<egp_tool_rerun>>": egp_tool_rerun,
            "<<egp_additional_mob>>": egp_additional_mob,
            "<<mfl_additional_mob>>": mfl_additional_mob,
            "<<personnel_additional_mob>>": personnel_additional_mob,
            "<<mfl_standby_rate>>": mfl_standby_rate,
            "<<egp_standby_rate>>": egp_standby_rate,
            "<<personnel_standby_rate>>": personnel_standby_rate,
        }

        for i in range(7):
            replacements[f"<<segment_{i+1}>>"] = f"{pipe_diameter}x{segment_lengths[i]}"
            replacements[f"<<price_segment_{i+1}>>"] = prices[i]

        # Replace placeholders
        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, val)

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
        st.success("✅ Document generated!")
        st.download_button(
            label="📥 Download Filled Document",
            data=buffer,
            file_name="Filled_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
