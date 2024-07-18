# Outlook Email Attachment Fetcher

## Description
This Python script automates the process of fetching PDF and DOCX attachments from specified senders in your Outlook inbox. It categorizes the attachments into predefined folders based on their filenames and handles potential filename conflicts.

## Features
- Securely stores and retrieves Outlook credentials using the system's keychain.
- Reads sender email addresses from a text file.
- Fetches both PDF and DOCX attachments.
- Categorizes attachments into folders: VitalsReports, DESRI DAILY, IRS1, and MISC.
- Handles filename conflicts by appending a counter to duplicate filenames.
- Provides detailed console output about the fetching process.

## Requirements
- Python 3.6+
- exchangelib
- keyring

## Installation
1. Clone this repository or download the script.
2. Install the required packages:
   ```sh
   pip install exchangelib keyring
   ```

## Setup
1. Create a file named `sender_emails.txt` in the same directory as the script.
2. Add the email addresses of the senders you want to fetch attachments from, one per line.

Example `sender_emails.txt`:
```
sender1@example.com
sender2@example.com
sender3@example.com
```

## Usage
1. Run the script:
   ```sh
   python outlook_attachment_fetcher.py
   ```
2. On the first run, you'll be prompted to enter your Outlook email address and password. These will be securely stored for future use.
3. Enter the number of days to look back for emails when prompted.
4. The script will fetch attachments and save them in the `FetchedAttachments` folder, organized into subfolders.

## Folder Structure
```
FetchedAttachments/
├── VitalsReports/
├── DESRI DAILY/
├── IRS1/
└── MISC/
```

## Error Handling
- If the `sender_emails.txt` file is not found or is empty, an error message will be displayed.
- If no messages are found from a sender, it will be reported in the console output.
- If no attachments are found from a sender, it will be explicitly stated in the output.

## Security Note
This script uses the `keyring` library to securely store your Outlook credentials in your system's keychain. Never share your credentials or include them directly in the script.

## Troubleshooting
- If you need to reset your stored credentials, you can use the `keyring` command-line tool or delete the stored credentials through your system's keychain manager.
- Ensure you have the necessary permissions to access the specified Outlook account.

---

