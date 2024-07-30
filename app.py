import streamlit as st
import pandas as pd
import json
from utils import load_excel, save_to_excel, generate_prompts
from openai_utils import get_email_sequence
from email_utils import send_emails
import logging

# Set the page configuration for a wider layout
st.set_page_config(layout="wide")

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'selected_email_number' not in st.session_state:
    st.session_state.selected_email_number = "Email 1"
if 'current_email_index' not in st.session_state:
    st.session_state.current_email_index = 0

# Add a logo to the app
logo_url = "Cvent_image.png"  # Replace with your image path or URL
st.image(logo_url, width=150)  # Adjust width as needed

# Add a headline for the app
st.title("Email Management System")

# Step 1: Upload Excel Document
if st.session_state.step == 1:
    st.header("Step 1: Upload Contacts..[Excel File]")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file:
        st.session_state.df = load_excel(uploaded_file)
        st.session_state.step = 2
        st.success("Document uploaded and data loaded successfully!")
        
        # Display the dataframe with all columns visible
        st.dataframe(st.session_state.df, use_container_width=True)

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
        st.dataframe(st.session_state.selected_df, use_container_width=True)
        
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
            st.dataframe(st.session_state.selected_df, use_container_width=True)

# Step 5: Generate Email Sequences
if st.session_state.step == 5:
    if 'selected_df' in st.session_state:
        st.header("Step 5: Generate Email Sequences")
        
        if st.button("Generate Sequences"):
            st.session_state.selected_df['email_sequence'] = st.session_state.selected_df['personalized_prompt'].apply(get_email_sequence)
            st.session_state.selected_email_number = "Email 1"  # Default value or adjust as needed
            st.session_state.step = 6
            st.success("Email sequences generated!")
            st.dataframe(st.session_state.selected_df, use_container_width=True)

# Step 6: Filter by Email Number
if st.session_state.step == 6:
    if 'selected_df' in st.session_state:
        st.header("Step 6: Filter by Email")
        email_numbers = ["Email 1", "Email 2", "Email 3"]
        selected_email_number = st.selectbox("Select Email Number", email_numbers)
        
        # Update the session state
        st.session_state.selected_email_number = selected_email_number
        
        def get_selected_email(email_sequence):
            try:
                if pd.isna(email_sequence):
                    return {}
                email_dict = json.loads(email_sequence)
                return email_dict.get(st.session_state.selected_email_number, {})
            except (TypeError, json.JSONDecodeError):
                return {}
        
        st.session_state.selected_df['selected_email'] = st.session_state.selected_df['email_sequence'].apply(get_selected_email)
        st.dataframe(st.session_state.selected_df[['contact name', 'selected_email']], use_container_width=True)
        
        if st.button("Proceed to Edit Emails"):
            st.session_state.step = 7
            
 # Step 7: Edit Emails
if st.session_state.step == 7:
    if 'selected_df' in st.session_state:
        st.header("Step 7: Edit Emails")

        # Get the selected email number
        selected_email_number = st.session_state.selected_email_number
        
        # Add a flag to check if multiple contacts are selected
        multiple_contacts_selected = len(st.session_state.selected_df) > 1

        for index, row in st.session_state.selected_df.iterrows():
            email_content = json.loads(row['email_sequence']).get(selected_email_number, {})
            if email_content:
                subject_line = email_content.get('Subject Line', '')
                body = email_content.get('Body', '')
                edited_email_key = f"edit_{index}"
                
                # Display the email content in a text area
                edited_email = st.text_area(
                    f"Edit Email for {row['contact name']} - {selected_email_number}",
                    value=f"Subject: {subject_line}\n\n{body}",
                    key=edited_email_key,
                    height = 700
                )
                
                # Save button specific to this email
                if st.button(f"Save Email {selected_email_number} for {row['contact name']}", key=f"save_{index}"):
                    # Extract subject and body from edited_email
                    lines = edited_email.split('\n\n', 1)
                    if len(lines) == 2:
                        st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0].replace('Subject: ', '')
                        st.session_state.selected_df.at[index, 'prepared_body'] = lines[1]
                    else:
                        st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0]
                        st.session_state.selected_df.at[index, 'prepared_body'] = ''
                    
                    st.success(f"Email for {row['contact name']} saved!")

        # Button to save all changes if multiple contacts are selected
        if multiple_contacts_selected:
            if st.button("Save All Changes"):
                for index, row in st.session_state.selected_df.iterrows():
                    email_content = json.loads(row['email_sequence']).get(selected_email_number, {})
                    if email_content:
                        subject_line = email_content.get('Subject Line', '')
                        body = email_content.get('Body', '')
                        edited_email = st.text_area(
                            f"Edit Email for {row['contact name']} - {selected_email_number}",
                            value=f"Subject: {subject_line}\n\n{body}",
                            key=f"edit_all_{index}"
                        )
                        # Extract subject and body from edited_email
                        lines = edited_email.split('\n\n', 1)
                        if len(lines) == 2:
                            st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0].replace('Subject: ', '')
                            st.session_state.selected_df.at[index, 'prepared_body'] = lines[1]
                        else:
                            st.session_state.selected_df.at[index, 'prepared_subject'] = lines[0]
                            st.session_state.selected_df.at[index, 'prepared_body'] = ''
                
                st.success("All changes saved!")

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