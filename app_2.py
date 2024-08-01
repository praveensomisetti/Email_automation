import streamlit as st
import pandas as pd
import json
from utils import load_excel, save_to_excel, generate_prompts
from openai_utils import get_email_sequence
from email_utils import send_emails

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'selected_email_number' not in st.session_state:
    st.session_state.selected_email_number = "Email 1"

# Step 1: Upload Excel Document
if st.session_state.step == 1:
    st.header("Step 1: Upload Excel Document")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file:
        st.session_state.df = load_excel(uploaded_file)
        st.session_state.step = 2
        st.success("Document uploaded and data loaded successfully!")
        st.dataframe(st.session_state.df)

# Step 2: Select Contacts
if st.session_state.step == 2:
    if 'df' in st.session_state:
        st.header("Step 2: Select Contacts")
        contact_names = st.session_state.df['contact name'].unique().tolist()
        contact_names.insert(0, "Select All")
        selected_contacts = st.multiselect("Select Contact Names", contact_names)
        
        if "Select All" in selected_contacts:
            selected_contacts = contact_names[1:]  # Exclude "Select All"
        
        st.session_state.selected_df = st.session_state.df[st.session_state.df['contact name'].isin(selected_contacts)]
        st.dataframe(st.session_state.selected_df)
        
        if st.button("Proceed to Generate Prompts"):
            st.session_state.step = 4

# Step 4: Generate Personalized Prompts
if st.session_state.step == 4:
    if 'selected_df' in st.session_state:
        st.header("Step 4: Generate Personalized Prompts")
        
        if st.button("Generate Prompts"):
            st.session_state.selected_df = generate_prompts(st.session_state.selected_df)
            st.session_state.step = 5
            st.success("Prompts generated!")
            st.dataframe(st.session_state.selected_df)

# Step 5: Generate Email Sequences
if st.session_state.step == 5:
    if 'selected_df' in st.session_state:
        st.header("Step 5: Generate Email Sequences")
        
        if st.button("Generate Sequences"):
            st.session_state.selected_df['email_sequence'] = st.session_state.selected_df['personalized_prompt'].apply(get_email_sequence)
            st.session_state.step = 6
            st.success("Email sequences generated!")
            st.dataframe(st.session_state.selected_df)

# Step 6: Filter by Email Number
if st.session_state.step == 6:
    if 'selected_df' in st.session_state:
        st.header("Step 6: Filter by Email Number")
        email_numbers = ["Email 1", "Email 2", "Email 3"]
        selected_email_number = st.selectbox("Select Email Number", email_numbers)
        
        # Update the session state
        st.session_state.selected_email_number = selected_email_number
        
        st.session_state.selected_df['selected_email'] = st.session_state.selected_df['email_sequence'].apply(
            lambda x: json.loads(x).get(st.session_state.selected_email_number)
        )
        st.dataframe(st.session_state.selected_df[['contact name', 'selected_email']])
        
        if st.button("Proceed to Edit Emails"):
            st.session_state.step = 7

# Step 7: Edit Emails
if st.session_state.step == 7:
    if 'selected_df' in st.session_state:
        st.header("Step 7: Edit Emails")
        for index, row in st.session_state.selected_df.iterrows():
            email_content = json.loads(row['email_sequence']).get(st.session_state.selected_email_number, {})
            if email_content:
                subject_line = email_content.get('Subject Line', '')
                body = email_content.get('Body', '')
                edited_email = st.text_area(f"Edit Email for {row['contact name']}", value=f"Subject: {subject_line}\n\n{body}", key=index)
                
                if st.button(f"Save Email {index}", key=f"save_{index}"):
                    # Extract subject and body from edited_email
                    lines = edited_email.split('\n\n', 1)
                    if len(lines) == 2:
                        st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0].replace('Subject: ', '')
                        st.session_state.selected_df.at[index, 'prepared_body'] = lines[1]
                    else:
                        st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0]
                        st.session_state.selected_df.at[index, 'prepared_body'] = ''
                    
                    st.success(f"Email for {row['contact name']} saved!")
        
        if st.button("Proceed to Send Emails"):
            st.session_state.step = 8

# Step 8: Send Emails
if st.session_state.step == 8:
    if 'selected_df' in st.session_state:
        st.header("Step 8: Send Emails")
        if st.button("Send Emails"):
            save_to_excel(st.session_state.selected_df, "final_emails.xlsx")
            results = send_emails(st.session_state.selected_df)
            st.success("All emails sent successfully!")
            for result in results:
                st.write(result)
                
                
                
                #
