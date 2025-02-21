import os
import tempfile
from docx import Document
import streamlit as st
from docx2pdf import convert
import platform

# Check if running on Windows or cloud
IS_WINDOWS = platform.system() == "Windows"

# Import Windows-specific modules only on Windows
if IS_WINDOWS:
    import pythoncom
    import win32com.client

def convert_to_pdf(input_docx, output_pdf):
    """Convert DOCX to PDF using win32com with explicit COM initialization"""
    if not IS_WINDOWS:
        st.warning("PDF conversion is only available in Windows. Downloading DOCX instead.")
        return False

    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        # Create Word application
        word = win32com.client.dynamic.Dispatch('Word.Application')
        word.Visible = False

        try:
            # Convert paths to absolute
            input_path = os.path.abspath(input_docx)
            output_path = os.path.abspath(output_pdf)
            
            # Open and convert
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # wdFormatPDF = 17
            doc.Close()
            
            return os.path.exists(output_pdf)
            
        finally:
            word.Quit()
            
    except Exception as e:
        st.error(f"PDF Conversion error: {str(e)}")
        return False
        
    finally:
        pythoncom.CoUninitialize()

# Define template paths
TEMPLATE_DIR = "templates"
template_paths = {
    "AI Automation": "Ai_automation.docx",
    "Digital Marketing": "DM Proposal.docx",
    "Business Automations": "Business Automations Proposal.docx",
    "IT Consultation": "Contract Agreement.docx"
}

def replace_text_preserve_formatting(doc, replacements):
    """Replace text while preserving formatting and images"""
    def replace_in_paragraph(paragraph, replacements):
        paragraph_text = paragraph.text
        runs = paragraph.runs
        
        for key, value in replacements.items():
            if key in paragraph_text:
                # Format price values if needed
                if "price" in key.lower() or "amount" in key.lower() or key == "{Additional}":
                    if isinstance(value, (int, float)):
                        value = f"$ {value:,.2f}"
                
                # Special handling for Additional Features text box
                if key == "{Additional}":
                    for i, run in enumerate(runs):
                        if "{Additional}" in run.text:
                            # Preserve any text before the placeholder
                            prefix = run.text[:run.text.find("{Additional}")]
                            run.text = prefix + str(value)
                            # Clear any remaining parts of the placeholder
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

    # Process text boxes first (they have priority)
    for shape in doc.inline_shapes:
        try:
            if shape._inline.graphic.graphicData.pic is not None:
                for element in shape._inline.graphic.graphicData.pic.xpath('.//w:txbxContent//w:p'):
                    try:
                        p = doc.paragraphs[0].__class__(element, doc.paragraphs[0]._parent)
                        if p.text.strip():
                            replace_in_paragraph(p, replacements)
                    except:
                        pass
        except:
            pass

    # Process all paragraphs including those in text boxes
    for element in doc._element.xpath('//w:txbxContent//w:p'):
        try:
            p = doc.paragraphs[0].__class__(element, doc.paragraphs[0]._parent)
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

    # Final pass for any remaining text boxes
    for element in doc._element.xpath('.//w:drawing//wp:anchor//w:txbxContent//w:p'):
        try:
            p = doc.paragraphs[0].__class__(element, doc.paragraphs[0]._parent)
            if p.text.strip():
                replace_in_paragraph(p, replacements)
        except:
            continue

def generate_proposal(proposal_type, client_name, replacements):
    """Generate proposal document with given replacements"""
    try:
        template_path = os.path.join(TEMPLATE_DIR, template_paths[proposal_type])
        doc = Document(template_path)
        replace_text_preserve_formatting(doc, replacements)
        
        temp_dir = tempfile.mkdtemp()
        output_docx = os.path.join(temp_dir, f"{proposal_type}_{client_name}.docx")
        output_pdf = os.path.join(temp_dir, f"{proposal_type}_{client_name}.pdf")
        
        # Save DOCX
        doc.save(output_docx)
        
        # Try PDF conversion
        try:
            convert(output_docx, output_pdf)
            with open(output_pdf, "rb") as pdf_file:
                pdf_bytes = pdf_file.read()
                st.success("PDF generated successfully!")
                return (pdf_bytes, 
                       f"{proposal_type}_{client_name}.pdf",
                       "application/pdf")
        except Exception as pdf_error:
            st.warning("PDF conversion failed, providing DOCX instead")
            with open(output_docx, "rb") as docx_file:
                docx_bytes = docx_file.read()
                return (docx_bytes,
                       f"{proposal_type}_{client_name}.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
    except Exception as e:
        st.error(f"Error generating proposal: {str(e)}")
        return None
    finally:
        try:
            if os.path.exists(temp_dir):
                import shutil
                shutil.rmtree(temp_dir)
        except:
            pass