import streamlit as st
import datetime
from pdf_generator import generate_proposal

def render_ba_form():
    st.header("Client Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", key="ba_name")
        contact_no = st.text_input("Contact Number", key="ba_contact")
        email_id = st.text_input("Email ID", key="ba_email_id")
    with col2:
        proposal_date = st.date_input("Date", datetime.datetime.now(), key="ba_date")
        validity_date = st.date_input("Validity Date", key="ba_validity")

    # Mutually Agreed Points
    mutually_agreed_points = st.text_area("Mutually Agreed Points", key="ba_points")
    
    # Week 1 Details
    st.header("Week 1 Details")
    week1_descrptn = st.text_area("Week 1 Description", key="ba_week1_desc")
    week1_price = st.number_input("Week 1 Price", min_value=0.0, step=0.01, format="%.2f", key="ba_week1_price")

    # Future Services Pricing
    st.header("Future Services Pricing")
    col1, col2 = st.columns(2)
    with col1:
        ai_auto_price = st.number_input("AI Automations Price", min_value=0.0, step=0.01, format="%.2f", key="ba_ai_auto")
        whts_price = st.number_input("WhatsApp Automation Price", min_value=0.0, step=0.01, format="%.2f", key="ba_whts")
        crm_price = st.number_input("CRM Setup Price", min_value=0.0, step=0.01, format="%.2f", key="ba_crm")
        email_price = st.number_input("Email Marketing Setup Price", min_value=0.0, step=0.01, format="%.2f", key="ba_email")
        make_price = st.number_input("Make/Zapier Automation Price", min_value=0.0, step=0.01, format="%.2f", key="ba_make")
    with col2:
        firefly_price = st.number_input("Firefly Meeting Price", min_value=0.0, step=0.01, format="%.2f", key="ba_firefly")
        chatbot_price = st.number_input("AI Chatbot Price", min_value=0.0, step=0.01, format="%.2f", key="ba_chatbot")
        pdf_gen_pr = st.number_input("PDF Generation Price", min_value=0.0, step=0.01, format="%.2f", key="ba_pdf")
        ai_mdl_price = st.number_input("AI Social Media Price", min_value=0.0, step=0.01, format="%.2f", key="ba_ai_mdl")
        cstm_ai_price = st.number_input("Custom AI Models Price", min_value=0.0, step=0.01, format="%.2f", key="ba_cstm_ai")

    replacements = {
        "{client_name}": client_name,
        "{contact_no}": contact_no,
        "{email_id}": email_id,
        "{date}": proposal_date.strftime("%d/%m/%Y"),
        "{validity_date}": validity_date.strftime("%d/%m/%Y"),
        "{mutually_agreed_points}": mutually_agreed_points,
        "{week1_descrptn}": week1_descrptn,
        "{week1_price}": f"$ {week1_price:,.2f}",
        "{ai_auto_price}": f"$ {ai_auto_price:,.2f}",
        "{whts_price}": f"$ {whts_price:,.2f}",
        "{crm_price}": f"$ {crm_price:,.2f}",
        "{email_price}": f"$ {email_price:,.2f}",
        "{make_price}": f"$ {make_price:,.2f}",
        "{firefly_price}": f"$ {firefly_price:,.2f}",
        "{chatbot_price}": f"$ {chatbot_price:,.2f}",
        "{pdf_gen_pr}": f"$ {pdf_gen_pr:,.2f}",
        "{ai_mdl_price}": f"$ {ai_mdl_price:,.2f}",
        "{cstm_ai_price}": f"$ {cstm_ai_price:,.2f}"
    }

    if st.button("Generate BA Proposal", key="ba_generate"):
        result = generate_proposal("Business Automations", client_name, replacements)
        if result:
            file_data, file_name, mime_type = result
            st.download_button(
                label=f"Download {file_name}",
                data=file_data,
                file_name=file_name,
                mime=mime_type
            ) 