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
st.title("Email Personalization")

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
        
        if st.button("Generate Prompts", key="generate_prompts_button"):
            st.session_state.step = 4

# Step 4: Generate Personalized Prompts
if st.session_state.step == 4:
    if 'selected_df' in st.session_state:
        st.header("Step 4: Generate Personalized Prompts")
        
        if st.button("Generate Prompts", key="generate_prompts_button_2"):
            st.session_state.selected_df = generate_prompts(st.session_state.selected_df)
            st.session_state.step = 5
            st.success("Prompts generated!")
            st.dataframe(st.session_state.selected_df, use_container_width=True)

# Step 5: Generate Email Sequences
if st.session_state.step == 5:
    if 'selected_df' in st.session_state:
        st.header("Step 5: Generate Emails")
        
        if st.button("Generate Emails", key="generate_emails_button"):
            st.session_state.selected_df['email_sequence'] = st.session_state.selected_df['personalized_prompt'].apply(get_email_sequence)
            st.session_state.selected_email_number = "Email 1"  # Default value or adjust as needed
            st.session_state.step = 6
            st.success("Email sequences generated!")
            st.dataframe(st.session_state.selected_df, use_container_width=True)

# Step 7: Edit Emails
def update_df():
    row_index = st.session_state.row_index
    email_data2 = json.loads(st.session_state.selected_df.at[row_index, st.session_state.col_name])
    email_data2["Email 1"]["Body"] = st.session_state.text_area
    st.session_state.selected_df.at[row_index, st.session_state.col_name] = json.dumps(email_data2)

if st.session_state.step == 6:
    if 'selected_df' in st.session_state:
        if 'row_index' not in st.session_state:
            st.session_state.row_index = 0
        if 'col_name' not in st.session_state:
            st.session_state.col_name = "email_sequence"
        if 'text_area' not in st.session_state:
            try:
                email_data = json.loads(st.session_state.selected_df.at[st.session_state.row_index, st.session_state.col_name])["Email 1"]
                st.session_state.text_area = email_data["Body"]
            except KeyError as e:
                st.error(f"KeyError: {e}")
            except Exception as e:
                st.error(f"Unexpected error: {e}")

        try:
            # Display current row's value in a text area
            email_data = json.loads(st.session_state.selected_df.at[st.session_state.row_index, st.session_state.col_name])["Email 1"]
            text_area_value = email_data["Body"]
            st.text_area("Edit value", value=text_area_value, key='text_area', on_change=update_df, height=700)
        except KeyError as e:
            st.error(f"KeyError: {e}")
        except Exception as e:
            st.error(f"Unexpected error: {e}")

        # Navigation buttons
        col1, col2 = st.columns(2)

        with col1:
            if st.button('Previous Email', key="previous_email_button"):
                if st.session_state.row_index > 0:
                    update_df()
                    st.session_state.row_index -= 1
                    st.rerun()
                else:
                    st.warning("Already at the first row.")

        with col2:
            if st.button('Next Email', key="next_email_button"):
                if st.session_state.row_index < len(st.session_state.selected_df) - 1:
                    update_df()
                    st.session_state.row_index += 1
                    st.rerun()
                else:
                    st.warning("Already at the last row.")

        if st.button("Save changes", key="save_changes_button"):
            update_df()
            st.session_state.step = 7

# Step 8: Send Emails
if st.session_state.step == 7:
    if 'selected_df' in st.session_state:
        st.header("Step 7: Send Emails")
        if st.button("Send Emails", key="send_emails_button"):
            save_to_excel(st.session_state.selected_df, "final_emails.xlsx")
            results = send_emails(st.session_state.selected_df)
            st.success("All emails sent successfully!")
            for result in results:
                st.write(result)
