# streamlit run streamlit_app.py

import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Word Template Generator", layout="centered")
st.title("📄 Word Template Generator")

# Select template type
template_type = st.radio("Select Document Type", ["ILI", "Quotation"])

# Upload template
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if template_file:
    with st.form("input_form"):
        st.write("📝 Fill the required fields")

        if template_type == "ILI":
            # ILI Template Fields
            st.subheader("🔹 General Info")
            project_name = st.text_input("Project Name")
            client_name = st.text_input("Client Name")
            contact_person = st.text_input("Contact Person")
            offer_number = st.text_input("Offer Number")
            quotation_date = st.date_input("Quotation Date")
            mobilization_days = st.text_input("Mobilization Days")
            inspection_days = st.text_input("Inspection Days")
            total_days = st.text_input("Total Duration (Days)")
            pipe_diameter = st.text_input("Pipe Diameter (e.g., 16inch)")

            st.subheader("🔹 Segment Lengths")
            segment_lengths = [st.text_input(f"Segment {i+1} Length") for i in range(7)]

            st.subheader("🔹 Segment Prices")
            prices = [st.text_input(f"Price for Segment {i+1}") for i in range(7)]

            st.subheader("🔹 Additional Costs")
            mfl_tool_rerun = st.text_input("MFL Tool Re-run Cost")
            egp_tool_rerun = st.text_input("EGP Tool Re-run Cost")
            egp_additional_mob = st.text_input("EGP Additional Mobilization")
            mfl_additional_mob = st.text_input("MFL/TFI Additional Mobilization")
            personnel_additional_mob = st.text_input("Personnel Additional Mobilization")
            mfl_standby_rate = st.text_input("MFL/TFI Standby Rate")
            egp_standby_rate = st.text_input("EGP Standby Rate")
            personnel_standby_rate = st.text_input("Personnel Standby Rate")

        elif template_type == "Quotation":
            # Quotation Template Fields
            st.subheader("🔹 Quotation Info")
            project_title = st.text_input("Project Title")
            project_name = st.text_input("Project Name")
            pipeline_size = st.text_input("Pipeline Size")
            client_name = st.text_input("Client Name")
            location = st.text_input("Location")
            mail_subject = st.text_input("Subject")
            service_type = st.text_input("Service Type")
            service_description = st.text_input("Service Description")
            uom_type = st.text_input("Unit of Measurement")
            price = st.text_input("Base Price")
            gst = st.text_input("GST Amount")
            total_price = st.text_input("Total Price")
            duration1 = st.text_input("Mobilization Duration")
            duration2 = st.text_input("Inspection Duration")
            duration3 = st.text_input("Total Duration")

            mfl_tool_rerun = st.text_input("MFL Tool Re-run Cost")
            egp_tool_rerun = st.text_input("EGP Tool Re-run Cost")
            egp_additional_mob = st.text_input("EGP Additional Mobilization")
            mfl_additional_mob = st.text_input("MFL/TFI Additional Mobilization")
            personnel_additional_mob = st.text_input("Personnel Additional Mobilization")
            mfl_standby_rate = st.text_input("MFL/TFI Standby Rate")
            egp_standby_rate = st.text_input("EGP Standby Rate")

        submitted = st.form_submit_button("Generate Document")

    if submitted:
        doc = Document(template_file)
        replacements = {}

        if template_type == "ILI":
            replacements = {
                "<<project_name>>": project_name,
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
                replacements[f"<<segment_{i+1}>>"] = segment_lengths[i]
                replacements[f"<<price_segment_{i+1}>>"] = prices[i]

        elif template_type == "Quotation":
            replacements = {
                "<<project_title>>": project_title,
                "<<project_name>>": project_name,
                "<<pipeline_size>>": pipeline_size,
                "<<client_name>>": client_name,
                "<<location>>": location,
                "<<mail_subject>>": mail_subject,
                "<<service_type>>": service_type,
                "<<service_description>>": service_description,
                "<<uom_type>>": uom_type,
                "<<price>>": price,
                "<<gst>>": gst,
                "<<total_price>>": total_price,
                "<<duration1>>": duration1,
                "<<duration2>>": duration2,
                "<<duration3>>": duration3,
                "<<mfl_tool_rerun>>": mfl_tool_rerun,
                "<<egp_tool_rerun>>": egp_tool_rerun,
                "<<egp_additional_mob>>": egp_additional_mob,
                "<<mfl_additional_mob>>": mfl_additional_mob,
                "<<personnel_additional_mob>>": personnel_additional_mob,
                "<<mfl_standby_rate>>": mfl_standby_rate,
                "<<egp_standby_rate>>": egp_standby_rate,
            }

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

        # Save and provide download
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.success("✅ Document generated!")
        st.download_button(
            label="📥 Download Word Document",
            data=buffer,
            file_name=f"{template_type}_Filled_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
