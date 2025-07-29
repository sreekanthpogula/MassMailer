import streamlit as st
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

# Flask context for Jinja2 rendering
mass_mailer = Flask(__name__)
mass_mailer.app_context().push()

# Configure Flask-Mail
mass_mailer.config['MAIL_SERVER'] = os.getenv('MAIL_SERVER')
mass_mailer.config['MAIL_PORT'] = int(os.getenv('MAIL_PORT'))
mass_mailer.config['MAIL_USE_TLS'] = os.getenv('MAIL_USE_TLS') == 'True'
mass_mailer.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
# mass_mailer.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
mass_mailer.config['MAIL_DEFAULT_SENDER'] = (
    os.getenv('MAIL_DEFAULT_SENDER_NAME'),
    os.getenv('MAIL_DEFAULT_SENDER_EMAIL')
)

mail = Mail(mass_mailer)

# Load default HTML template
def load_template():
    with open('templates/email_template.html', 'r', encoding='utf-8') as f:
        return f.read()

# Validate data and return structured errors
def validate_excel_data(df):
    error_map = {}

    for index, row in df.iterrows():
        issues = []
        row_number = index + 2
        associate_name = str(row.get("AssociateName", "")).strip()
        associate_id = str(row.get("AssociateID", "")).strip()
        to_email = str(row.get("Associate Email", "")).strip()
        cl_email = str(row.get("CL Email", "")).strip()
        pm_email = str(row.get("PM Email", "")).strip()

        if not to_email or "@senecaglobal.com" not in to_email:
            issues.append(("Associate Email", f"Invalid Associate Email – {to_email}", to_email))

        if not cl_email or "@senecaglobal.com" not in cl_email:
            issues.append(("CL Email", f"Invalid CL Email – {cl_email}", cl_email))

        if not pm_email or "@senecaglobal.com" not in pm_email:
            issues.append(("PM Email", f"Invalid PM Email – {pm_email}", pm_email))

        if "N" not in associate_id:
            issues.append(("AssociateID", f"Invalid Associate ID – {associate_id}", associate_id))

        if len(associate_name.split()) < 2:
            issues.append(("AssociateName", f"AssociateName should be 'Firstname Lastname' – {associate_name}", associate_name))

        kra_file_name = f"{associate_name}.pdf"
        kra_path = os.path.join("kra_files", kra_file_name)
        if not os.path.exists(kra_path):
            issues.append(("KRA File", f"Missing KRA file for {associate_name}", ""))

        if issues:
            error_map[index] = issues

    return error_map

# Retry wrapper for sending email
def send_email_with_retry(msg, retries=3, delay=3):
    for attempt in range(1, retries + 1):
        try:
            mail.send(msg)
            return True
        except Exception as e:
            logging.warning(f"Attempt {attempt} failed: {e}")
            time.sleep(delay)
    return False

# Main email sending function
def send_bulk_emails(dry_run=False, df=None, template=None):
    logs = []

    if df is None:
        df = pd.read_excel('excel_files/KRA.xlsx')
    if template is None:
        template = load_template()

    for _, row in df.iterrows():
        associate_name = row['AssociateName']
        to_email = row['Associate Email']
        cc_emails = [row['CL Email'], row['PM Email']]
        kra_file_name = f"{associate_name}.pdf"
        kra_path = os.path.join('kra_files', kra_file_name)

        if not os.path.exists(kra_path):
            msg = f"Missing KRA file: {kra_path}"
            logging.warning(msg)
            logs.append(msg)
            continue

        html_body = render_template_string(template, associate=associate_name)

        msg = Message(
            subject='Your KRA Document',
            recipients=[to_email],
            cc=cc_emails,
            html=html_body
        )
        with open(kra_path, 'rb') as f:
            msg.attach(kra_file_name, 'application/pdf', f.read())

        if dry_run:
            log = f"[DRY-RUN] Would send to {associate_name} ({to_email}), CC: {cc_emails}"
            logging.info(log)
            logs.append(log)
        else:
            success = send_email_with_retry(msg)
            if success:
                log = f"[SENT] Email sent to {associate_name} ({to_email})"
                logging.info(log)
                logs.append(log)
            else:
                log = f"[FAILED] Could not send email to {associate_name}"
                logging.error(log)
                logs.append(log)

    return logs

# --- Streamlit UI ---
st.title("📬 Mass Mailer application")
st.subheader("Email Template Selection")

template_options = ["Default Template"]
template_map = {"Default Template": """
    <font face='Arial' size='3'>
    <body>
    Dear {{ associate }},
    <br><br>
    IGNORE MY EARLIER EMAIL.
    <br><br>
    We have defined goals and objectives for client delivery performance to achieve continual improvement...
    <br><br>
    All the best!!
    <br>
    </body>
    </font>
"""
}

uploaded_templates = st.file_uploader("Upload Custom Templates", type=["html", "txt"], accept_multiple_files=True)
if uploaded_templates:
    for file in uploaded_templates:
        content = file.read().decode("utf-8")
        template_map[file.name] = content
        template_options.append(file.name)

col1, col2 = st.columns(2)

with col1:
    if "Default Template" in template_map:
        st.markdown("**Sample Email Template**")
        st.download_button(
            label="📥 Sample Template",
            data=template_map["Default Template"],
            file_name="Sample_Template.html",
            mime="text/html"
        )

# Define file path
kra_file_path = "massmailer/excel_files/KRA.xlsx"
kra_map = {"Default KRA Template":"""
AssociateID, AssociateName, PM Email, CL Email, Associate Email
N1070, Sreekanth Pogula, sreekanth.pogula@senecaglobal.com, sreekanth.pogula@senecaglobal.com, sreekanth.pogula@senecaglobal.com
"""}

with col2:
    if "Default KRA Template" in kra_map:
        st.markdown("**Sample KRA Excel File**")
        st.download_button(
            label="📥 Sample KRA File",
            data=kra_map["Default KRA Template"],
            file_name="KRA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


selected_template = st.selectbox("Choose a Template", template_options)
template_input = st.text_area("Email HTML Template", value=template_map[selected_template], height=100)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Helper: highlight invalid cells
def highlight_invalid_cells(df, error_map):
    def apply_styles(row):
        styles = [""] * len(row)
        if row.name in error_map:
            col_indices = {col: i for i, col in enumerate(df.columns)}
            for col, _, _ in error_map[row.name]:
                if col in col_indices:
                    styles[col_indices[col]] = "background-color: #FFCCCC"
        return styles
    return df.style.apply(apply_styles, axis=1)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.dataframe(df)

    if st.button("Dry Run"):
        if not template_input.strip():
            st.error("Please provide a valid email template.")
        else:
            errors = validate_excel_data(df)
            if errors:
                st.warning("Validation Errors Found:")
                styled_df = highlight_invalid_cells(df, errors)
                st.dataframe(styled_df)

                for row_idx, issues in errors.items():
                    st.markdown(f"### Row {row_idx + 2}")
                    for col, msg, suggestion in issues:
                        new_value = st.text_input(f"{msg}", value=suggestion, key=f"{row_idx}_{col}")
                        if new_value.strip():
                            df.at[row_idx, col] = new_value.strip()
            else:
                logs = send_bulk_emails(dry_run=True, df=df, template=template_input)
                st.success("Dry-run completed.")
                st.code("\n".join(logs))

    if st.button("Send Emails"):
        if not template_input.strip():
            st.error("Please provide a valid email template.")
        else:
            with st.spinner("Sending emails..."):
                logs = send_bulk_emails(dry_run=False, df=df, template=template_input)
                st.success("Emails sent.")

# --- CLI Support ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send KRA emails")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    send_bulk_emails(dry_run=args.dry_run)