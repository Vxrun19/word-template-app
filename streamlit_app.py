import streamlit as st
from docx import Document
from datetime import datetime
import io

st.set_page_config(page_title="Quotation Generator", layout="centered")

st.title("📄 Quotation Generator App")

# --- Select Template Type ---
template_type = st.radio("Select Document Type", ["ILI", "Quotation", "ILI Offer"])

# --- ILI Template ---
if template_type == "ILI":
    st.subheader("🔹 ILI Quotation Details")
    project_name = st.text_input("Project Name")
    client_name = st.text_input("Client Name")
    quotation_date = st.date_input("Quotation Date")
    quotation_number = st.text_input("Quotation Number")
    tool = st.text_input("Tool")
    od = st.text_input("OD")
    pipeline_length = st.text_input("Pipeline Length")
    schedule = st.text_input("Schedule")
    job_location = st.text_input("Job Location")

# --- Quotation Template ---
elif template_type == "Quotation":
    st.subheader("🔹 Quotation Template Info")
    client_name = st.text_input("Client Name")
    quotation_date = st.date_input("Quotation Date")
    quotation_number = st.text_input("Quotation Number")
    attention = st.text_input("Attention")
    subject = st.text_input("Subject")
    scope_of_work = st.text_area("Scope of Work")
    location = st.text_input("Location")
    offer_validity = st.text_input("Offer Validity")
    commercial_terms = st.text_area("Commercial Terms")

# --- ILI Offer Template ---
elif template_type == "ILI Offer":
    st.subheader("🔹 General Info")
    project_name = st.text_input("Project Name")
    client_name = st.text_input("Client Name")
    quotation_date = st.date_input("Quotation Date")
    mobilization_days = st.text_input("Mobilization Days")
    inspection_days = st.text_input("Inspection Days")
    total_days = st.text_input("Total Duration (Days)")

    st.subheader("🔹 Segments Info")
    segment_lengths = [st.text_input(f"Segment {i+1} Length") for i in range(7)]
    prices = [st.text_input(f"Price for Segment {i+1}") for i in range(7)]

    st.subheader("🔹 Financial Info")
    total_price = st.text_input("Total Price")
    gst = st.text_input("GST (%)")
    grand_total = st.text_input("Grand Total")

# --- Submit Button ---
submitted = st.button("Generate Document")

# --- Generate and Download Document ---
if submitted:
    if template_type == "ILI":
        doc = Document("Quotation Format (1).docx")
        replacements = {
            "<<project_name>>": project_name,
            "<<client_name>>": client_name,
            "<<quotation_date>>": quotation_date.strftime("%d/%m/%Y"),
            "<<quotation_number>>": quotation_number,
            "<<tool>>": tool,
            "<<od>>": od,
            "<<pipeline_length>>": pipeline_length,
            "<<schedule>>": schedule,
            "<<job_location>>": job_location,
        }

    elif template_type == "Quotation":
        doc = Document("Quotation Format (1).docx")
        replacements = {
            "<<client_name>>": client_name,
            "<<quotation_date>>": quotation_date.strftime("%d/%m/%Y"),
            "<<quotation_number>>": quotation_number,
            "<<attention>>": attention,
            "<<subject>>": subject,
            "<<scope_of_work>>": scope_of_work,
            "<<location>>": location,
            "<<offer_validity>>": offer_validity,
            "<<commercial_terms>>": commercial_terms,
        }

    elif template_type == "ILI Offer":
        doc = Document("Quotation Format (1).docx")
        replacements = {
            "<<project_name>>": project_name,
            "<<client_name>>": client_name,
            "<<quotation_date>>": quotation_date.strftime("%d/%m/%Y"),
            "<<mobilization_days>>": mobilization_days,
            "<<inspection_days>>": inspection_days,
            "<<total_days>>": total_days,
            "<<total_price>>": total_price,
            "<<gst>>": gst,
            "<<grand_total>>": grand_total,
        }
        for i in range(7):
            replacements[f"<<segment_{i+1}_length>>"] = segment_lengths[i]
            replacements[f"<<segment_{i+1}_price>>"] = prices[i]

    # Replace text in document
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

    # Save to in-memory buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # File download name
    filename = f"{template_type}_Quotation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

    # Download Button
    st.success("✅ Document ready!")
    st.download_button(
        label="📥 Download Document",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
