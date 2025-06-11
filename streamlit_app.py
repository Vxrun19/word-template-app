import streamlit as st
from docx import Document
import datetime
from io import BytesIO
import os
import base64
import tempfile
import pypandoc
import uuid

def fill_placeholders(doc: Document, replacements: dict):
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                for run in p.runs:
                    run.text = run.text.replace(key, val)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)

def generate_offer(template_path, replacements):
    doc = Document(template_path)
    fill_placeholders(doc, replacements)

    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)

    tmp_id = str(uuid.uuid4())
    docx_path = os.path.join(tempfile.gettempdir(), f"{tmp_id}.docx")
    pdf_path = os.path.join(tempfile.gettempdir(), f"{tmp_id}.pdf")

    with open(docx_path, "wb") as f:
        f.write(docx_buffer.read())

    pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)

    return docx_path, pdf_path

def embed_pdf(file_path):
    with open(file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode("utf-8")
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600px" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

# --- Streamlit UI ---
st.title("Offer Letter Generator")

mode = st.selectbox("Select Offer Type", ["ILI", "Quotation", "Old ILI"])

with st.form("offer_form"):
    st.header("Project & Client Info")
    project_name = st.text_input("Project Name")
    client_name = st.text_input("Client Name")
    contact_person = st.text_input("Contact Person")
    offer_number = st.text_input("Offer Number")
    quotation_date = st.date_input("Quotation Date", value=datetime.date.today())

    replacements = {
        "<<project_name>>": project_name,
        "<<client_name>>": client_name,
        "<<contact_person>>": contact_person,
        "<<offer_number>>": offer_number,
        "<<quotation_date>>": quotation_date.strftime("%d-%m-%Y"),
    }

    if mode == "ILI":
        st.header("ILI Details")
        mobilization_days = st.number_input("Mobilization Days", min_value=0, step=1)
        inspection_days = st.number_input("Inspection Days", min_value=0, step=1)
        pipe_diameter = st.text_input("Pipeline Diameter")
        for i in range(1, 8):
            replacements[f"<<segment_{i}>>"] = st.text_input(f"Segment {i} Description", key=f"seg{i}")
            replacements[f"<<price_segment_{i}>>"] = st.text_input(f"Segment {i} Price", key=f"price{i}")
        replacements.update({
            "<<mobilization_days>>": str(mobilization_days),
            "<<inspection_days>>": str(inspection_days),
            "<<total_days>>": str(mobilization_days + inspection_days),
            "<<pipe_diameter>>": pipe_diameter,
            "<<mfl_tool_rerun>>": st.text_input("MFL Tool Rerun Charge"),
            "<<egp_tool_rerun>>": st.text_input("EGP Tool Rerun Charge"),
            "<<egp_additional_mob>>": st.text_input("EGP Additional Mobilization"),
            "<<mfl_additional_mob>>": st.text_input("MFL Additional Mobilization"),
            "<<personnel_additional_mob>>": st.text_input("Personnel Additional Mobilization"),
            "<<mfl_standby_rate>>": st.text_input("MFL Standby Rate per Day"),
            "<<egp_standby_rate>>": st.text_input("EGP Standby Rate per Day"),
            "<<personnel_standby_rate>>": st.text_input("Personnel Standby Rate per Day")
        })
        template = "quotation_template_ili.docx"

    elif mode == "Quotation":
        st.header("Quotation Segments")
        for i in range(1, 8):
            replacements[f"<<segment_{i}>>"] = st.text_input(f"Segment {i} Description", key=f"qseg{i}")
            replacements[f"<<price_segment_{i}>>"] = st.text_input(f"Segment {i} Price", key=f"qprice{i}")
        template = "quotation_template_quotation.docx"

    else:  # Old ILI
        st.header("Old ILI Details")
        inspection_scope = st.text_area("Inspection Scope")
        for i in range(1, 8):
            replacements[f"<<segment_{i}>>"] = st.text_input(f"Segment {i} Description", key=f"oldseg{i}")
            replacements[f"<<price_segment_{i}>>"] = st.text_input(f"Segment {i} Price", key=f"oldprice{i}")
        replacements["<<inspection_scope>>"] = inspection_scope
        template = "quotation_template_old_ili.docx"

    submitted = st.form_submit_button("Generate Offer")

if submitted:
    docx_path, pdf_path = generate_offer(template, replacements)
    st.success("Offer generated successfully!")

    with open(docx_path, "rb") as f:
        st.download_button("Download Word File", f, f"{offer_number}_{mode.replace(' ', '_')}.docx")

    embed_pdf(pdf_path)
