from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment, EWSDateTime, EWSTimeZone
import os
from datetime import datetime, timedelta
import keyring
import getpass

def get_or_create_credentials():
    service_name = 'OutlookEmailFetcher'
    email_account = keyring.get_password(service_name, 'email')
    password = keyring.get_password(service_name, 'password')
    
    if not email_account or not password:
        email_account = input("Enter your Outlook email address: ")
        password = getpass.getpass("Enter your Outlook password: ")
        
        keyring.set_password(service_name, 'email', email_account)
        keyring.set_password(service_name, 'password', password)
        
        print("Credentials saved securely.")
    
    return email_account, password

def read_sender_emails(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Sender email file not found: {file_path}")
        return []

    with open(file_path, 'r') as f:
        return [line.strip() for line in f if line.strip()]

def get_folder_name(file_name):
    if file_name.startswith("VitalsReports-Daily"):
        return "VitalsReports"
    elif file_name.startswith("DESRI Daily Executive Summary Report"):
        return "DESRI DAILY"
    elif file_name.startswith("IRS1 Daily Report"):
        return "IRS1"
    else:
        return "MISC"

def fetch_attachments_outlook():
    email_address, password = get_or_create_credentials()
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sender_file = os.path.join(script_dir, "sender_emails.txt")
    sender_emails = read_sender_emails(sender_file)
    
    if not sender_emails:
        print("Error: No sender emails found. Please add sender email addresses to the file and run the script again.")
        return []

    days_ago = int(input("Enter the number of days to look back: "))

    tz = EWSTimeZone.localzone()
    date = EWSDateTime.now(tz=tz) - timedelta(days=days_ago)

    base_folder = os.path.join(script_dir, "FetchedAttachments")
    folders = ["VitalsReports", "DESRI DAILY", "IRS1", "MISC"]
    for folder in folders:
        folder_path = os.path.join(base_folder, folder)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

    fetched_files = []

    try:
        print("Connecting to Outlook server...")
        credentials = Credentials(email_address, password)
        config = Configuration(server='outlook.office365.com', credentials=credentials)
        account = Account(primary_smtp_address=email_address, config=config, autodiscover=False, access_type=DELEGATE)
        print("Connection successful")

        for sender in sender_emails:
            print(f"Searching for messages from {sender} since {date}...")
            messages = account.inbox.filter(datetime_received__gt=date, sender__contains=sender)
            message_list = list(messages)
            
            if not message_list:
                print(f"No messages found from {sender}")
                continue

            print(f"Found {len(message_list)} messages from {sender}")
            sender_attachments = 0

            for message in message_list:
                for attachment in message.attachments:
                    if isinstance(attachment, FileAttachment) and (attachment.name.lower().endswith('.pdf') or attachment.name.lower().endswith('.docx')):
                        folder_name = get_folder_name(attachment.name)
                        save_folder = os.path.join(base_folder, folder_name)
                        save_path = os.path.join(save_folder, attachment.name)
                        
                        # Ensure unique filename
                        counter = 1
                        while os.path.exists(save_path):
                            name, ext = os.path.splitext(attachment.name)
                            save_path = os.path.join(save_folder, f"{name}_{counter}{ext}")
                            counter += 1
                        
                        with open(save_path, 'wb') as f:
                            f.write(attachment.content)
                        fetched_files.append(save_path)
                        sender_attachments += 1

            if sender_attachments == 0:
                print(f"0 attachments fetched from {sender}")
            else:
                print(f"{sender_attachments} attachments fetched from {sender}")

        return fetched_files

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        if hasattr(e, 'args') and len(e.args) > 1:
            print(f"Additional error details: {e.args}")
        return fetched_files

# Run the function
fetched_files = fetch_attachments_outlook()
print(f"\nTotal attachments retrieved: {len(fetched_files)}")
for file in fetched_files:
    print(file)