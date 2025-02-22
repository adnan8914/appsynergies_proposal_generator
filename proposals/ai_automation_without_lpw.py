import streamlit as st
import datetime
from pdf_generator import generate_proposal

def render_ai_automation_without_lpw_form():
    st.header("AI Automation without LPW")
    
    # Client Information
    st.subheader("Client Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name")
        email = st.text_input("Email")
        phone = st.text_input("Phone")
    with col2:    
        country = st.text_input("Country")
        proposal_date = st.date_input("Proposal Date")
        validity_date = st.date_input("Validity Date", 
                                    value=proposal_date + datetime.timedelta(days=365),
                                    min_value=proposal_date,
                                    help="Proposal validity end date")

    # Project Pricing
    st.subheader("Project Pricing")
    col1, col2 = st.columns(2)
    with col1:
        ai_calling_price = st.number_input("AI Calling", min_value=0.0, step=0.01)
        crm_price = st.number_input("CRM Automations", min_value=0.0, step=0.01)
    with col2:
        manychat_price = st.number_input("ManyChat & Make Automation", min_value=0.0, step=0.01)
        additional_price = st.number_input("Additional Features & Enhancements", min_value=0.0, step=0.01)

    # Calculate totals
    total_price = ai_calling_price + crm_price + manychat_price
    annual_maintenance = total_price * 0.20

    # Display totals
    st.subheader(f"Total Amount: ${total_price:,.2f}")
    st.subheader(f"Annual Maintenance: ${annual_maintenance:,.2f}")

    if st.button("Generate Proposal"):
        if not client_name:
            st.error("Please enter client name")
            return
            
        replacements = {
            "{client_name}": client_name,
            "{Email_address}": email,
            "{Phone_no}": phone,
            "{country_name}": country,
            "{date}": proposal_date.strftime("%d/%m/%Y"),
            "{validity_date}": validity_date.strftime("%d/%m/%Y"),
            "{AI calling price}": f"$ {ai_calling_price:,.2f}",
            "{CRM Automation price}": f"$ {crm_price:,.2f}",
            "{Manychat price}": f"$ {manychat_price:,.2f}",
            "{Total amount}": f"$ {total_price:,.2f}",
            "{AM price}": f"$ {annual_maintenance:,.2f}",
            "{Additional}": f"$ {additional_price:,.2f}"
        }
        
        result = generate_proposal("AI Automation without LPW", client_name, replacements)
        
        if result:
            file_data, file_name, mime_type = result
            st.download_button(
                label="Download Proposal",
                data=file_data,
                file_name=file_name,
                mime=mime_type
            ) 