import streamlit as st
from proposals.ai_automation import render_ai_automation_form
from proposals.digital_marketing import render_dm_form
from proposals.business_automation import render_ba_form
from proposals.contract import render_contract_form
from proposals.ai_automation_without_lpw import render_ai_automation_without_lpw_form

def main():
    st.title("Proposal Generator")
    
    proposal_type = st.radio(
        "Select Proposal Type",
        ["AI Automation",
         "AI Automation without LPW",
         "Digital Marketing", 
         "Business Automations", 
         "IT Consultation"]
    )

    if proposal_type == "AI Automation":
        render_ai_automation_form()
    elif proposal_type == "AI Automation without LPW":
        render_ai_automation_without_lpw_form()
    elif proposal_type == "Digital Marketing":
        render_dm_form()
    elif proposal_type == "Business Automations":
        render_ba_form()
    else:  # IT Consultation
        render_contract_form()

if __name__ == "__main__":
    main() 