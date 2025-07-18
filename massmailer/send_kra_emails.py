import pandas as pd
import os
import time
from flask import Flask, render_template_string
from flask_mail import Mail, Message
from dotenv import load_dotenv
import logging
import argparse

# Load environment variables
load_dotenv()

# Logging setup
os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    filename='logs/email_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Initialize Flask mass_mailer context for Jinja2 rendering
mass_mailer = Flask(__name__)
mass_mailer.app_context().push()

# Flask-Mail configuration
mass_mailer.config['MAIL_SERVER'] = os.getenv('MAIL_SERVER')
mass_mailer.config['MAIL_PORT'] = int(os.getenv('MAIL_PORT'))
mass_mailer.config['MAIL_USE_TLS'] = os.getenv('MAIL_USE_TLS') == 'True'
mass_mailer.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
mass_mailer.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
mass_mailer.config['MAIL_DEFAULT_SENDER'] = (
    os.getenv('MAIL_DEFAULT_SENDER_NAME'),
    os.getenv('MAIL_DEFAULT_SENDER_EMAIL')
)

mail = Mail(mass_mailer)

# Load email template
def load_template():
    with open('templates/email_template.html', 'r') as f:
        return f.read()

# Send email with retry logic
def send_email_with_retry(msg, retries=3, delay=3):
    for attempt in range(1, retries + 1):
        try:
            mail.send(msg)
            return True
        except Exception as e:
            logging.warning(f"Attempt {attempt} failed: {e}")
            time.sleep(delay)
    return False

# Main function
def send_bulk_emails(dry_run=False):
    try:
        df = pd.read_excel('excel_files/KRA.xlsx')
        template = load_template()

        for index, row in df.iterrows():
            associate_name = row['AssociateName']
            to_email = row['Associate Email']
            cc_emails = [row['CL Email'], row['PM Email']]
            
            # Auto-generate file name from Associate Name
            kra_file_name = f"{associate_name.replace(' ', ' ')}.pdf"
            kra_path = os.path.join('kra_files', kra_file_name)

            if not os.path.exists(kra_path):
                msg = f"Missing KRA file for {associate_name}: {kra_path}"
                print(msg)
                logging.warning(msg)
                continue

            html_body = render_template_string(template, name=associate_name)

            msg = Message(
                subject='Your KRA Document',
                recipients=[to_email],
                cc=cc_emails,
                html=html_body
            )

            with open(kra_path, 'rb') as f:
                msg.attach(kra_file_name, 'application/pdf', f.read())

            if dry_run:
                print(f"[DRY-RUN] Would send to {associate_name} ({to_email}), CC: {cc_emails}, File: {kra_file_name}")
                logging.info(f"[DRY-RUN] {associate_name} | {to_email} | {cc_emails} | {kra_file_name}")
                continue

            success = send_email_with_retry(msg)
            if success:
                log_msg = f"[SENT] Email sent to {associate_name} ({to_email})"
                print(log_msg)
                logging.info(log_msg)
            else:
                err_msg = f"[FAILED] Could not send email to {associate_name} after retries."
                print(err_msg)
                logging.error(err_msg)


        print("✅ All emails processed.")
        logging.info("All emails processed.")

    except Exception as e:
        print(f"[FATAL ERROR] {e}")
        logging.critical(f"Fatal error: {e}")

# CLI Argument parsing
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Send bulk KRA emails.')
    parser.add_argument('--dry-run', action='store_true', help='Preview emails without sending.')
    args = parser.parse_args()

    send_bulk_emails(dry_run=args.dry_run)
