import streamlit as st
import datetime
from pdf_generator import generate_proposal

def render_contract_form():
    st.header("Contract Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", key="contract_name")
        client_company_address = st.text_area("Company Address", key="contract_address")
    with col2:
        contract_date = st.date_input("Contract Date", datetime.datetime.now(), key="contract_date")

    replacements = {
        "{date}": contract_date.strftime("%d/%m/%Y"),
        "{client_name}": client_name,
        "{client_company_address}": client_company_address
    }

    if st.button("Generate Contract", key="contract_generate"):
        result = generate_proposal("IT Consultation", client_name, replacements)
        if result:
            file_data, file_name, mime_type = result
            st.download_button(
                label=f"Download {file_name}",
                data=file_data,
                file_name=file_name,
                mime=mime_type
            ) 