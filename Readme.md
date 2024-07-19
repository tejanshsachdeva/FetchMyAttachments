# Fetch My Attachments

This Streamlit app allows users to fetch email attachments from their Outlook account based on specific criteria. The attachments will be organized into a zip file with categorized folders.

## Installation

To run this app locally, follow these steps:

1. Clone the repository:

   ```bash
   git clone https://github.com/tejanshsachdeva/FetchMyAttachments.git
   cd FetchMyAttachments
   ```
2. Install the required packages:

   ```bash
   pip install -r requirements.txt
   ```
3. Run the Streamlit app:

   ```bash
   streamlit run script.py
   ```

## Usage

1. **Enter Your Outlook Email Address and Password:**

   - **Email Address:** Input the Outlook email address where you receive the emails from which attachments are to be fetched.
   - **Password:** Enter the password for your Outlook email. Click on the eye icon to reveal the password if needed. Press 'Enter' or click outside the box to proceed after entering the password.
2. **Upload Sender Email Addresses:**

   - **Option 1: Upload a Text File:** Click on the 'Browse files' button or drag and drop a text file (.txt) containing the sender email addresses. Each email address should be on a new line.
   - **Option 2: Enter Manually:** Alternatively, you can manually enter the sender email addresses in the provided text box. Each email address should be on a new line.
3. **Specify the Date Range:**

   - **Start Date:** Select the start date from which you want to search for attachments.
   - **End Date:** Select the end date up to which you want to search for attachments.
4. **Fetch Attachments:**

   - Click the 'Fetch Attachments' button to start the process. The app will access your Outlook account, search for emails from the specified senders within the provided date range, and fetch the attachments.
5. **Download the Zip File:**

   - Once the process is complete, the app will generate a zip file containing the fetched attachments organized into the following folders:
     ```
     FetchedAttachments/
     ├── VitalsReports/
     ├── DESRI DAILY/
     ├── IRS1/
     └── MISC/
     ```

## Features

- **Email Authentication:** Securely connect to your Outlook email using your credentials.
- **Sender Filtering:** Fetch attachments only from specified senders.
- **Date Range Filtering:** Specify a date range to limit the search for attachments.
- **Categorized Folders:** Automatically organize fetched attachments into predefined folders based on their filenames.
- **Download as Zip:** Easily download all fetched attachments in a single zip file.

## Troubleshooting and Support

By following these instructions, you can seamlessly fetch and categorize attachments from your Outlook emails. Should you encounter any problems or need further assistance, please consult the error messages provided by the app for guidance. For additional support, feel free to create an issue on our GitHub repository: [Submit an Issue](https://github.com/tejanshsachdeva/FetchMyAttachments/issues).

## Walkthrough

For a detailed step-by-step guide on using the app, please refer to the [walkthrough.md](https://github.com/tejanshsachdeva/FetchMyAttachments/blob/main/WalkThrough.md) file in the repository.

## Contributing

Feel free to contribute by cloning this repo. There is a [test folder](https://github.com/tejanshsachdeva/FetchMyAttachments/tree/main/test) in this repo which has `frontend.py` and `backend.py` code. This uses Flask as the backend to send requests to the Outlook API. Contribute to it, improve the security measures, and feel free to submit a pull request.

---
