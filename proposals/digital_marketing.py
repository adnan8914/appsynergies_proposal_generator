import streamlit as st
import datetime
from pdf_generator import generate_proposal

def render_dm_form():
    st.header("Client Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", key="dm_name")
        designation = st.text_input("Designation", key="dm_designation")
        contact_no = st.text_input("Contact Number", key="dm_contact")
    with col2:
        email_id = st.text_input("Email ID", key="dm_email")
        proposal_date = st.date_input("Date", datetime.datetime.now(), key="dm_date")
    
    # Mutually Agreed Points
    mutually_agreed_points = st.text_area("Mutually Agreed Points")
    
    # Digital Marketing Services Pricing
    st.header("Digital Marketing Services Pricing")
    col1, col2 = st.columns(2)
    with col1:
        social_media_posts = st.number_input("30 Creative Social Media Posts", min_value=0.0, step=0.01, format="%.2f", key="dm_smp")
        research_and_dev = st.number_input("Marketing Research + 1 Month Ads", min_value=0.0, step=0.01, format="%.2f", key="dm_rnd")
    with col2:
        monthly_cost = st.number_input("Monthly Maintenance", min_value=0.0, step=0.01, format="%.2f", key="dm_monthly")
    
    # Calculate totals
    subtotal = social_media_posts + research_and_dev + monthly_cost
    gst = subtotal * 0.18  # 18% GST
    total_amount = subtotal + gst
    advance = total_amount * 0.5
    balance = total_amount * 0.5
    
    replacements = {
        "{client_name}": client_name,
        "{designation}": designation,
        "{contact_no}": contact_no,
        "{email_id}": email_id,
        "{date}": proposal_date.strftime("%d/%m/%Y"),
        "{Mutually_agreed_points}": mutually_agreed_points,
        "{3d_SMP}": f"$ {social_media_posts:,.2f}",
        "{R&D}": f"$ {research_and_dev:,.2f}",
        "{monthly_cost}": f"$ {monthly_cost:,.2f}",
        "{gst}": f"$ {gst:,.2f}",
        "{total_amount}": f"$ {total_amount:,.2f}",
        "{Advance}": f"$ {advance:,.2f}",
        "{balance}": f"$ {balance:,.2f}"
    }

    if st.button("Generate DM Proposal", key="dm_generate"):
        result = generate_proposal("Digital Marketing", client_name, replacements)
        if result:
            file_data, file_name, mime_type = result
            st.download_button(
                label=f"Download {file_name}",
                data=file_data,
                file_name=file_name,
                mime=mime_type
            ) 