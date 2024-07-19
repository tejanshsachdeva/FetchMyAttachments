import streamlit as st
from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment, EWSDateTime, EWSTimeZone
import os
from datetime import datetime, timedelta
import zipfile
import io
import re

# Function to get credentials securely using session state
def get_credentials():
    if "email_account" not in st.session_state:
        st.session_state.email_account = ""
    if "password" not in st.session_state:
        st.session_state.password = ""
    
    email_account = st.text_input("Enter your Outlook email address:", value=st.session_state.email_account)
    password = st.text_input("Enter your Outlook password:", type="password")
    
    if email_account and password:
        st.session_state.email_account = email_account
        st.session_state.password = password
    
    return email_account, password

def read_sender_emails(file_content):
    return [line.strip() for line in file_content.split('\n') if line.strip()]

def get_folder_name(file_name):
    if file_name.startswith("VitalsReports-Daily"):
        return "VitalsReports"
    elif file_name.startswith("DESRI Daily Executive Summary Report"):
        return "DESRI DAILY"
    elif file_name.startswith("IRS1 Daily Report"):
        return "IRS1"
    else:
        return "MISC"

def fetch_attachments_outlook(email_address, password, sender_emails, start_date, end_date):
    if not email_address or not password:
        st.error("Please enter your credentials.")
        return None

    if not sender_emails:
        st.error("No sender emails provided. Please upload a file with sender email addresses.")
        return None

    # Validate email format
    email_regex = re.compile(r"[^@]+@[^@]+\.[^@]+")
    if not email_regex.match(email_address):
        st.error("Invalid email format.")
        return None

    tz = EWSTimeZone.localzone()
    start_date = EWSDateTime.from_datetime(datetime.combine(start_date, datetime.min.time()).replace(tzinfo=tz))
    end_date = EWSDateTime.from_datetime(datetime.combine(end_date, datetime.max.time()).replace(tzinfo=tz))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
        try:
            with st.spinner("Connecting to Outlook server..."):
                credentials = Credentials(email_address, password)
                config = Configuration(server='outlook.office365.com', credentials=credentials)
                account = Account(primary_smtp_address=email_address, config=config, autodiscover=False, access_type=DELEGATE)
            st.success("Connection successful")

            progress_bar = st.progress(0)
            status_text = st.empty()

            total_attachments = 0

            for i, sender in enumerate(sender_emails):
                status_text.text(f"Searching for messages from {sender} between {start_date.date()} and {end_date.date()}...")
                messages = account.inbox.filter(datetime_received__range=(start_date, end_date), sender__contains=sender)
                message_list = list(messages)
                
                if not message_list:
                    st.warning(f"No messages found from {sender}")
                    continue

                st.info(f"Found {len(message_list)} messages from {sender}")
                sender_attachments = 0

                for message in message_list:
                    for attachment in message.attachments:
                        if isinstance(attachment, FileAttachment) and (attachment.name.lower().endswith('.pdf') or attachment.name.lower().endswith('.docx')):
                            folder_name = get_folder_name(attachment.name)
                            # Prevent path traversal attacks
                            safe_folder_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', folder_name)
                            safe_file_name = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', attachment.name)
                            file_path = os.path.join(safe_folder_name, safe_file_name)
                            
                            zip_file.writestr(file_path, attachment.content)
                            sender_attachments += 1
                            total_attachments += 1

                if sender_attachments == 0:
                    st.warning(f"0 attachments fetched from {sender}")
                else:
                    st.success(f"{sender_attachments} attachments fetched from {sender}")

                progress_bar.progress((i + 1) / len(sender_emails))

            status_text.text("Attachment fetching completed.")
            return zip_buffer, total_attachments

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            return None

st.title("Fetch My Attachments")

email_address, password = get_credentials()

uploaded_file = st.file_uploader("Upload a file with sender email addresses (one per line)", type="txt")
sender_emails = []

if uploaded_file is not None:
    sender_emails = read_sender_emails(uploaded_file.getvalue().decode())
    st.write(f"Loaded {len(sender_emails)} sender email(s)")

else:
    sender_emails_input = st.text_area("Enter sender email addresses (one per line)")
    if sender_emails_input:
        sender_emails = read_sender_emails(sender_emails_input)
        st.write(f"Loaded {len(sender_emails)} sender email(s)")

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", datetime.now() - timedelta(days=7))
with col2:
    end_date = st.date_input("End Date", datetime.now())

if start_date > end_date:
    st.error("Error: End date must fall after start date.")

if st.button("Fetch Attachments"):
    if (uploaded_file is not None or sender_emails) and start_date <= end_date:
        result = fetch_attachments_outlook(email_address, st.session_state.password, sender_emails, start_date, end_date)
        if result:
            zip_buffer, total_attachments = result
            st.success(f"Total attachments retrieved: {total_attachments}")
            
            # Offer the zip file for download
            zip_buffer.seek(0)
            st.download_button(
                label="Download Attachments as ZIP",
                data=zip_buffer,
                file_name="fetched_attachments.zip",
                mime="application/zip"
            )
    else:
        st.error("Please upload a file with sender email addresses or enter them manually, and ensure the end date is after the start date.")
