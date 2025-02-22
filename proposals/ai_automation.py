import streamlit as st
import datetime
from pdf_generator import generate_proposal

def render_ai_automation_form():
    st.header("AI Automation with Landing Page Proposal")
    
    # Client Information
    st.subheader("Client Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", key="ai_name")
        email = st.text_input("Email", key="ai_email")
        phone = st.text_input("Phone", key="ai_phone")
    with col2:    
        country = st.text_input("Country", key="ai_country")
        proposal_date = st.date_input("Proposal Date", key="ai_date")
        validity_date = st.date_input("Validity Date", 
                                    value=proposal_date + datetime.timedelta(days=365),  # Default 1 year validity
                                    min_value=proposal_date,
                                    help="Proposal validity end date")
    
    # Project Pricing
    st.header("Project Pricing")
    col1, col2 = st.columns(2)
    with col1:
        landing_page_price = st.number_input("Landing Page Website", min_value=0.0, step=0.01, format="%.2f")
        admin_panel_price = st.number_input("Admin Panel", min_value=0.0, step=0.01, format="%.2f")
        crm_price = st.number_input("CRM Automations", min_value=0.0, step=0.01, format="%.2f")
    with col2:
        manychat_price = st.number_input("ManyChat & Make Automation", min_value=0.0, step=0.01, format="%.2f")
        social_media_price = st.number_input("Social Media Automation", min_value=0.0, step=0.01, format="%.2f")
        ai_calling_price = st.number_input("AI Calling", min_value=0.0, step=0.01, format="%.2f")
    
    additional_features_price = st.number_input(
        "Additional Features & Enhancements (USD per week)", 
        min_value=0.0, 
        step=0.01, 
        format="%.2f",
        value=250.00
    )
    
    total_price = (landing_page_price + admin_panel_price + crm_price + 
                  manychat_price + social_media_price + ai_calling_price)
    annual_maintenance = total_price * 0.20
    
    st.header("Signature Details")
    company_representative = st.text_input("Company Representative")

    replacements = {
        "{client_name}": client_name,
        "{Email_address}": email,
        "{Phone_no}": phone,
        "{country_name}": country,
        "{date}": proposal_date.strftime("%d/%m/%Y"),
        "{validity_date}": validity_date.strftime("%d/%m/%Y"),
        "{landing page price}": landing_page_price,
        "{admin panel price}": admin_panel_price,
        "{CRM Automation price}": crm_price,
        "{Manychat price}": manychat_price,
        "{SMP price}": social_media_price,
        "{AI calling price}": ai_calling_price,
        "{Total amount}": total_price,
        "{AM price}": annual_maintenance,
        "{Additional}": additional_features_price,
        "{company_representative}": company_representative,
    }

    if st.button("Generate Proposal"):
        if not client_name:
            st.error("Please enter client name")
            return
            
        result = generate_proposal("AI Automation with Landing Page", client_name, replacements)
        if result:
            file_data, file_name, mime_type = result
            st.download_button(
                label="Download Proposal",
                data=file_data,
                file_name=file_name,
                mime=mime_type
            ) 