import streamlit as st
import requests
from datetime import datetime, timedelta

API_URL = "http://localhost:5000"  # Replace with your actual backend URL when deploying

st.title("Fetch My Attachments")

email = st.text_input("Enter your Outlook email address:")
password = st.text_input("Enter your Outlook password:", type="password")

uploaded_file = st.file_uploader("Upload a file with sender email addresses (one per line)", type="txt")
sender_emails = []

if uploaded_file is not None:
    sender_emails = [line.strip() for line in uploaded_file.getvalue().decode().split('\n') if line.strip()]
    st.write(f"Loaded {len(sender_emails)} sender email(s)")
else:
    sender_emails_input = st.text_area("Enter sender email addresses (one per line)")
    if sender_emails_input:
        sender_emails = [line.strip() for line in sender_emails_input.split('\n') if line.strip()]
        st.write(f"Loaded {len(sender_emails)} sender email(s)")

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", datetime.now() - timedelta(days=7))
with col2:
    end_date = st.date_input("End Date", datetime.now())

if start_date > end_date:
    st.error("Error: End date must fall after start date.")

if st.button("Fetch Attachments"):
    if email and password and sender_emails and start_date <= end_date:
        # First, test the connection
        with st.spinner("Testing connection..."):
            response = requests.post(f"{API_URL}/test_connection", json={
                "email": email,
                "password": password
            })
            
            if response.status_code != 200:
                st.error(f"Connection failed: {response.json().get('error', 'Unknown error occurred')}")
            else:
                st.success("Connection successful! Fetching attachments...")
                
                # If connection is successful, proceed with fetching attachments
                with st.spinner("Fetching attachments..."):
                    response = requests.post(f"{API_URL}/fetch_attachments", json={
                        "email": email,
                        "password": password,
                        "sender_emails": sender_emails,
                        "start_date": start_date.isoformat(),
                        "end_date": end_date.isoformat()
                    }, stream=True)
                    
                    if response.status_code == 200:
                        st.success("Attachments fetched successfully!")
                        
                        # Offer the zip file for download
                        st.download_button(
                            label="Download Attachments as ZIP",
                            data=response.content,
                            file_name="fetched_attachments.zip",
                            mime="application/zip"
                        )
                    elif response.status_code == 404:
                        st.warning("No attachments found for the given criteria.")
                    else:
                        st.error(f"Error fetching attachments: {response.text}")
    else:
        st.error("Please fill in all required fields and ensure the end date is after the start date.")