from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment, EWSDateTime, EWSTimeZone
from datetime import datetime, timedelta
import zipfile
import io
import os
import logging

app = Flask(__name__)
CORS(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def get_folder_name(file_name):
    if file_name.startswith("VitalsReports-Daily"):
        return "VitalsReports"
    elif file_name.startswith("DESRI Daily Executive Summary Report"):
        return "DESRI DAILY"
    elif file_name.startswith("IRS1 Daily Report"):
        return "IRS1"
    else:
        return "MISC"

@app.route('/test_connection', methods=['POST'])
def test_connection():
    data = request.json
    email = data.get('email')
    password = data.get('password')

    if not email or not password:
        return jsonify({"error": "Email and password are required"}), 400

    try:
        credentials = Credentials(email, password)
        config = Configuration(server='outlook.office365.com', credentials=credentials)
        account = Account(primary_smtp_address=email, config=config, autodiscover=False, access_type=DELEGATE)
        
        # If we get here, the connection was successful
        return jsonify({"success": True, "message": "Connection successful"})

    except Exception as e:
        logger.error(f"Connection test failed: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/fetch_attachments', methods=['POST'])
def fetch_attachments():
    data = request.json
    email = data.get('email')
    password = data.get('password')
    sender_emails = data.get('sender_emails')
    start_date = datetime.strptime(data.get('start_date'), '%Y-%m-%d').date()
    end_date = datetime.strptime(data.get('end_date'), '%Y-%m-%d').date()

    if not all([email, password, sender_emails, start_date, end_date]):
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        credentials = Credentials(email, password)
        config = Configuration(server='outlook.office365.com', credentials=credentials)
        account = Account(primary_smtp_address=email, config=config, autodiscover=False, access_type=DELEGATE)

        tz = EWSTimeZone.localzone()
        start_date = EWSDateTime.from_datetime(datetime.combine(start_date, datetime.min.time()).replace(tzinfo=tz))
        end_date = EWSDateTime.from_datetime(datetime.combine(end_date, datetime.max.time()).replace(tzinfo=tz))

        zip_buffer = io.BytesIO()
        total_attachments = 0

        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for sender in sender_emails:
                logger.info(f"Searching for messages from {sender}")
                messages = account.inbox.filter(datetime_received__range=(start_date, end_date), sender__contains=sender)
                message_count = 0
                
                for message in messages:
                    message_count += 1
                    logger.info(f"Processing message: {message.subject}")
                    for attachment in message.attachments:
                        if isinstance(attachment, FileAttachment) and (attachment.name.lower().endswith('.pdf') or attachment.name.lower().endswith('.docx')):
                            folder_name = get_folder_name(attachment.name)
                            file_path = os.path.join(folder_name, attachment.name)
                            logger.info(f"Adding attachment: {file_path}")
                            zip_file.writestr(file_path, attachment.content)
                            total_attachments += 1

                logger.info(f"Processed {message_count} messages from {sender}")

        if total_attachments == 0:
            logger.warning("No attachments found")
            return jsonify({"error": "No attachments found"}), 404

        logger.info(f"Total attachments added to ZIP: {total_attachments}")
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='fetched_attachments.zip'
        )

    except Exception as e:
        logger.error(f"Error in fetch_attachments: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)