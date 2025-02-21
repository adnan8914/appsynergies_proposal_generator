import streamlit as st
import datetime
import tempfile
import os
from docx import Document
import platform

# Check if running on Windows or cloud
IS_WINDOWS = platform.system() == "Windows"

if IS_WINDOWS:
    import win32com.client
    import pythoncom
    from docx2pdf import convert
else:
    # Alternative PDF conversion for cloud (could offer DOCX download only)
    def convert(input_docx, output_pdf):
        st.warning("PDF conversion is only available in Windows. Downloading DOCX instead.")
        return False

def convert(input_docx, output_pdf):
    if IS_WINDOWS:
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch('Word.Application')
            doc = None
            try:
                word.Visible = False
                input_path = os.path.abspath(input_docx)
                output_path = os.path.abspath(output_pdf)
                doc = word.Documents.Open(input_path)
                doc.SaveAs(FileName=output_path, FileFormat=17)
                return True
            finally:
                if doc:
                    doc.Close()
                word.Quit()
        except Exception as e:
            st.error(f"Conversion error: {str(e)}")
            return False
        finally:
            pythoncom.CoUninitialize()
    return False

# Define template paths relative to the script
TEMPLATE_DIR = "templates"  # Create this directory in your repository
AI_TEMPLATE = os.path.join(TEMPLATE_DIR, "Ai_automation.docx")
DM_TEMPLATE = os.path.join(TEMPLATE_DIR, "DM Proposal.docx")
BA_TEMPLATE = os.path.join(TEMPLATE_DIR, "Business Automations Proposal.docx")
CONTRACT_TEMPLATE = os.path.join(TEMPLATE_DIR, "Contract Agreement.docx")

def replace_text_preserve_formatting(doc, replacements):
    """Replace text while preserving formatting and images"""
    def replace_in_paragraph(paragraph, replacements):
        paragraph_text = paragraph.text
        runs = paragraph.runs
        
        # Debug print to see what we're dealing with
        if "Additional Features" in paragraph_text:
            st.write("Found Additional Features paragraph:")
            st.write("Runs:", [r.text for r in runs])
        
        for key, value in replacements.items():
            if key in paragraph_text:
                # Format price values if needed
                if "price" in key.lower() or "amount" in key.lower() or key == "{Additional}":
                    if isinstance(value, (int, float)):
                        value = f"$ {value:,.2f}"
                
                # Special handling for Additional Features text box
                if "Additional Features" in paragraph_text and "Enhancements" in paragraph_text:
                    # Try to find the exact run with the placeholder
                    for i, run in enumerate(runs):
                        if "{Additional}" in run.text:
                            # Count tabs/spaces before the placeholder
                            prefix = run.text[:run.text.find("{Additional}")]
                            run.text = prefix + str(value)
                            # Clear any remaining parts
                            for j in range(i + 1, len(runs)):
                                if "{" in runs[j].text or "}" in runs[j].text:
                                    runs[j].text = ""
                            return
                
                # Normal replacement logic for other fields
                key_runs = []
                start_index = -1
                current_key_part = ""
                
                for i, run in enumerate(runs):
                    if not run.text:
                        continue
                    
                    current_key_part += run.text
                    if key in current_key_part:
                        for j in range(start_index if start_index >= 0 else i, i + 1):
                            key_runs.append(runs[j])
                        start_index = -1
                        current_key_part = ""
                    elif any(part in current_key_part for part in ["{", key[1:len(key)-1]]):
                        if start_index < 0:
                            start_index = i
                    else:
                        start_index = -1
                        current_key_part = ""
                
                if key_runs:
                    key_runs[0].text = str(value)
                    for run in key_runs[1:]:
                    
                        run.text = ""

    # Process all paragraphs including those in text boxes
    for paragraph in doc.element.xpath('//w:p'):
        try:
            p = doc.paragraphs[0].__class__(paragraph, doc.paragraphs[0]._parent)
            if p.text.strip():
                replace_in_paragraph(p, replacements)
        except:
            pass

    # Process regular paragraphs
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            replace_in_paragraph(paragraph, replacements)

    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip():
                        replace_in_paragraph(paragraph, replacements)

    # Process text boxes using Word XML
    for element in doc._element.xpath('.//w:drawing//wp:anchor//w:txbxContent//w:p'):
        try:
            text = element.text
            for key, value in replacements.items():
                if key in text:
                    if "price" in key.lower() or "amount" in key.lower() or key == "{Additional}":
                        if isinstance(value, (int, float)):
                            value = f"$ {value:,.2f}"
                    element.text = text.replace(key, str(value))
        except:
            continue

    # Try to process shapes directly
    try:
        for shape in doc.inline_shapes:
            if shape._inline.graphic.graphicData.pic is not None:
                txbox = shape._inline.graphic.graphicData.pic.nvPicPr.cNvPr.txBox
                if txbox is not None:
                    for paragraph in txbox.paragraphs:
                        replace_in_paragraph(paragraph, replacements)
    except:
        pass

def main():
    st.title("Proposal Generator")
    
    # Proposal Type Selection
    proposal_type = st.selectbox(
        "Select Proposal Type",
        ["AI Automation Proposal", "Digital Marketing Proposal", "Business Automations Proposal", "IT Consultation Contract"]
    )
    
    if proposal_type == "AI Automation Proposal":
        template_path = AI_TEMPLATE
    elif proposal_type == "Digital Marketing Proposal":
        template_path = DM_TEMPLATE
    elif proposal_type == "Business Automations Proposal":
        template_path = BA_TEMPLATE
    else:  # IT Consultation Contract
        template_path = CONTRACT_TEMPLATE
    
    # Client Information
    st.header("Client Information")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", key="ai_name")
        email = st.text_input("Email", key="ai_email")
        phone = st.text_input("Phone", key="ai_phone")
    with col2:    
        country = st.text_input("Country", key="ai_country")
        proposal_date = st.date_input("Date", datetime.datetime.now(), key="ai_date")
    
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
    
    # Additional Features & Enhancements
    additional_features_price = st.number_input("Additional Features & Enhancements (USD per week)", 
                                              min_value=0.0, 
                                              step=0.01, 
                                              format="%.2f",
                                              value=250.00)  # Default value set to 250
    
    total_price = (landing_page_price + admin_panel_price + crm_price + 
                  manychat_price + social_media_price + ai_calling_price)
    annual_maintenance = total_price * 0.20
    
    # Signature Details
    st.header("Signature Details")
    company_representative = st.text_input("Company Representative")

    # Updated replacements dictionary
    replacements = {
        "{client_name}": client_name,
        "{Email_address}": email,
        "{Phone_no}": phone,
        "{country_name}": country,
        "{date}": proposal_date.strftime("%d/%m/%Y"),
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

    # Generate proposal button and processing - same for all types
    if st.button("Generate Proposal"):
        if proposal_type == "AI Automation Proposal":
            if not all([client_name, email, phone, country]):
                st.error("Please fill in all required fields")
                return
        elif proposal_type == "Digital Marketing Proposal":
            if not all([client_name, designation, contact_no, email_id]):
                st.error("Please fill in all required fields")
                return
        elif proposal_type == "Business Automations Proposal":
            if not all([client_name, contact_no, email_id]):
                st.error("Please fill in all required fields")
                return
        else:  # IT Consultation Contract
            if not all([client_name, client_company_address]):
                st.error("Please fill in all required fields")
                return
            
        try:
            doc = Document(template_path)
            replace_text_preserve_formatting(doc, replacements)
            
            temp_dir = tempfile.mkdtemp()
            output_docx = os.path.join(temp_dir, f"{proposal_type}_{client_name}.docx")
            doc.save(output_docx)
            
            if IS_WINDOWS:
                # Try PDF conversion on Windows
                output_pdf = os.path.join(temp_dir, f"{proposal_type}_{client_name}.pdf")
                if convert(output_docx, output_pdf) and os.path.exists(output_pdf):
                    with open(output_pdf, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                        st.download_button(
                            label="Download Proposal PDF",
                            data=pdf_bytes,
                            file_name=f"{proposal_type}_{client_name}.pdf",
                            mime="application/pdf"
                        )
                    st.success("Proposal generated successfully!")
            else:
                # On cloud, offer DOCX download
                with open(output_docx, "rb") as docx_file:
                    docx_bytes = docx_file.read()
                    st.download_button(
                        label="Download Proposal DOCX",
                        data=docx_bytes,
                        file_name=f"{proposal_type}_{client_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("Proposal generated successfully!")
            
        except Exception as e:
            st.error(f"Error generating proposal: {str(e)}")
        finally:
            try:
                os.rmdir(temp_dir)
            except:
                pass

if __name__ == "__main__":
    main()