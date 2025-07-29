import streamlit as st
import pandas as pd
import os
import logging
from flask import Flask, render_template_string
from dotenv import load_dotenv
# from auth.auth import send_email_graph

# Load environment variables
load_dotenv()

# Logging setup
os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    filename='logs/email_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Flask app for Jinja2 rendering
mass_mailer = Flask(__name__)
mass_mailer.app_context().push()

# --- Email validation ---
def validate_excel_data(df):
    errors = []
    for index, row in df.iterrows():
        associate_name = row['AssociateName']
        to_email = row['Associate Email']
        cc_emails = [row['CL Email'], row['PM Email']]

        if pd.isna(to_email) or '@' not in to_email:
            errors.append(f"Row {index + 2}: Invalid or missing Associate Email for {associate_name}")

        kra_file_name = f"{associate_name.replace(' ', ' ')}.pdf"
        kra_path = os.path.join('kra_files', kra_file_name)
        if not os.path.exists(kra_path):
            errors.append(f"Row {index + 2}: Missing KRA file for {associate_name}: {kra_file_name}")

    return errors

# --- Email sending logic using Graph API ---
def send_bulk_emails_graph(df, template, dry_run=False):
    results = []

    for index, row in df.iterrows():
        associate_name = row['AssociateName']
        to_email = row['Associate Email']
        cc_emails = [row['CL Email'], row['PM Email']]
        kra_file_name = f"{associate_name.replace(' ', ' ')}.pdf"
        kra_path = os.path.join('kra_files', kra_file_name)

        if not os.path.exists(kra_path):
            results.append(f"[SKIPPED] {associate_name} – Missing file: {kra_file_name}")
            continue

        html_body = render_template_string(template, name=associate_name)

        if dry_run:
            dry_msg = f"[DRY-RUN] Would send to {associate_name} ({to_email}), CC: {cc_emails}, File: {kra_file_name}"
            results.append(dry_msg)
            logging.info(dry_msg)
            continue

        try:
            success = send_email_graph(to_email, cc_emails, subject="Your KRA Document", html_body=html_body)
            if success:
                success_msg = f"[SENT] Email sent to {associate_name} ({to_email})"
                results.append(success_msg)
            else:
                raise Exception("Microsoft Graph API failed.")
        except Exception as e:
            error_msg = f"[FAILED] {associate_name} – Error: {e}"
            results.append(error_msg)
            logging.error(error_msg)

    return results

# --- Streamlit UI ---
st.title("📬 Mass Mailer application")

# --- Template Selection ---
st.subheader("Email Template Selection")

template_options = ["Default Template"]
uploaded_templates = st.file_uploader("Upload Custom Templates", type=["txt", "html"], accept_multiple_files=True)

template_map = {
    "Default Template": """
<font face='Arial' size='3'>
<body>
Dear {{ associate }},
<br><br>

IGNORE MY EARLIER EMAIL.
<br><br>

We have defined goals and objectives for client delivery performance to achieve continual improvement and they are aligned to the key result areas (KRAs) of associates. The KRAs have been defined to improve work performance, teaming, collaboration and competence development of the associates. Please find attached a letter containing the description of KRAs applicable to you for the year 2025-26.
<br><br>

Please discuss your KRAs with your Reporting Manager and implement appropriate action plan for improving your work performance and competence. Your performance against the KRAs will be monitored by your Reporting Manager and reviewed periodically by the Associate Development Review Committee. HR department will be in touch with you and provide necessary guidance for your work performance and competency development planning, implementation, and development reviews.
<br><br>

All the best!!
<br>
</body>
</font>
"""
}

if uploaded_templates:
    for file in uploaded_templates:
        content = file.read().decode("utf-8")
        template_map[file.name] = content
        template_options.append(file.name)

selected_template = st.selectbox("Choose a Template", template_options)

# Populate the editor with the selected template
template_input = st.text_area("Email HTML Template", value=template_map[selected_template], height=300)


uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.dataframe(df)

    if st.button("Dry Run", key="dry_run"):
        if not template_input.strip():
            st.error("Please provide a valid email template.")
        else:
            errors = validate_excel_data(df)
            if errors:
                st.warning("Validation Errors:")
                st.code("\n".join(errors))
            else:
                dry_run_logs = send_bulk_emails_graph(df, template_input, dry_run=True)
                st.success("Dry-run completed successfully.")
                st.code("\n".join(dry_run_logs))

    if st.button("Send Emails", key="send_emails"):
        if not template_input.strip():
            st.error("Please provide a valid email template.")
        else:
            with st.spinner("Sending emails..."):
                results = send_bulk_emails_graph(df, template_input, dry_run=False)
                st.success("Email sending completed.")
                st.code("\n".join(results))

    
