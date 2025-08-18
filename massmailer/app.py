import smtplib
import logging
import os
import glob
import time
import pandas as pd
import streamlit as st
from jinja2 import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv
from datetime import datetime

# LLM imports
from langchain.chains import RetrievalQA
from langchain_community.embeddings import OllamaEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_community.document_loaders import PyPDFDirectoryLoader
from langchain_community.llms import ollama

# Load environment variables
load_dotenv()

config = {
    'MAIL_SERVER': os.getenv('MAIL_SERVER'),
    'MAIL_PORT': int(os.getenv('MAIL_PORT', 25)),
    'MAIL_USERNAME': os.getenv('MAIL_USERNAME'),
}

# Logging setup
os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    filename='logs/email_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Load default HTML template
def load_template():
    with open('templates/email_template.html', 'r', encoding='utf-8') as f:
        return f.read()


# Get the year range for the KRA document like "2023-24"
def get_year_range():
    current_year = datetime.now().year
    next_year_short = str(current_year + 1)[-2:]  # e.g., "26"
    return f"{current_year}-{next_year_short}"


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
        kra_path = os.path.join(".temp/kra_files/pdf_files", kra_file_name)
        if not os.path.exists(kra_path):
            issues.append(("KRA File", f"Missing KRA file for {associate_name}", ""))

        if issues:
            error_map[index] = issues

    return error_map


# Send a single email with retry logic
def send_email_with_retry(subject, to_email, cc_emails, html_body, attachment_path, config, retries=1, delay=0.5):
    for attempt in range(1, retries + 2):
        try:
            msg = MIMEMultipart()
            msg['From'] = config['MAIL_USERNAME']
            msg['To'] = to_email
            msg['Cc'] = ", ".join(cc_emails)
            msg['Subject'] = subject

            msg.attach(MIMEText(html_body, 'html'))

            if os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as f:
                    part = MIMEApplication(f.read(), _subtype='pdf')
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
                    

            smtp_server = smtplib.SMTP(config['MAIL_SERVER'], int(config['MAIL_PORT']))
            
            # smtp_server.login(config['MAIL_USERNAME'], config['MAIL_PASSWORD'])
            smtp_server.sendmail(
                from_addr=config['MAIL_USERNAME'],
                to_addrs=[to_email] + cc_emails,
                msg=msg.as_string()
            )
            smtp_server.quit()
            return True

        except Exception as e:
            logging.warning(f"Attempt {attempt} failed for {to_email}: {e}")
            time.sleep(delay)

    return False


def send_bulk_emails(dry_run=False, df=None, template=None):
    logs = []

    # Check for missing config
    for key, val in config.items():
        if not val:
            raise ValueError(f"Missing email config: {key} is not set in .env")

    if df is None:
        df = pd.read_excel('excel_files/KRA.xlsx')
    if template is None:
        template = load_template()

    for _, row in df.iterrows():
        associate_name = row.get("AssociateName", "").strip()
        to_email = row.get("Associate Email", "").strip()
        cc_emails = [row.get("CL Email", "").strip(), row.get("PM Email", "").strip()]
        kra_file_name = f"{associate_name}.pdf"
        kra_path = os.path.join('.temp/kra_files/pdf_files', kra_file_name)

        # Render HTML using Jinja2
        html_body = Template(template).render(associate=associate_name, year_range=get_year_range())

        if dry_run:
            log = f"[DRY-RUN] Able to send to {associate_name} with {to_email} | CC: {cc_emails}"
            logging.info(log)
            logs.append(log)
            continue

        success = send_email_with_retry(
            subject='Your KRA Document',
            to_email=to_email,
            cc_emails=cc_emails,
            html_body=html_body,
            attachment_path=kra_path,
            config=config
        )

        if success:
            log = f"[SENT] Email sent to {associate_name} with {to_email} | CC: {cc_emails}"
            logging.info(log)
        else:
            log = f"[FAILED] Could not send email to {associate_name} with {to_email}"
            logging.error(log)
        logs.append(log)

    return logs


# --- Streamlit UI ---
st.title("📬 Mass Mailer application")
st.subheader("Add Email Template")

template_options = ["Default Template"]
template_map = {"Default Template": """
<font face="Arial" size="3">
<body>
Dear {{ associate }},
<BR><BR>

IGNORE MY EARLIER EMAIL.
<BR><BR>

We have defined goals and objectives for client delivery performance to achieve continual improvement and they are aligned to the key result areas (KRAs) of associates. The KRAs have been defined to improve work performance, teaming, collaboration and competence development of the associates. Please find attached a letter containing the description of KRAs applicable to you for the year 2025-26.
<BR><BR>

Please discuss your KRAs with your Reporting Manager and implement appropriate action plan for improving your work performance and competence. Your performance against the KRAs will be monitored by your Reporting Manager and reviewed periodically by the Associate Development Review Committee. HR department will be in touch with you and provide necessary guidance for your work performance and competency development planning, implementation, and development reviews.
<BR><BR>

All the best!!
<BR>
</body>
</font>
"""
}

# uploaded_templates = st.file_uploader("Upload Custom Templates", type=["html", "txt"], accept_multiple_files=True)
# if uploaded_templates:
#     for file in uploaded_templates:
#         content = file.read().decode("utf-8")
#         template_map[file.name] = content
#         template_options.append(file.name)

# col1, col2 = st.columns(2)

# with col1:
#     if "Default Template" in template_map:
#         st.markdown("**Sample Email Template**")
#         st.download_button(
#             label="📥 Sample Template",
#             data=template_map["Default Template"],
#             file_name="Sample_Template.html",
#             mime="text/html"
#         )

# Define file path
# kra_file_path = "massmailer/excel_files/KRA.xlsx"
# kra_map = {"Default KRA Template":"""
# AssociateID, AssociateName, PM Email, CL Email, Associate Email
# N1070, Sreekanth Pogula, sreekanth.pogula@senecaglobal.com, sreekanth.pogula@senecaglobal.com, sreekanth.pogula@senecaglobal.com
# """}

# with col2:
#     if "Default KRA Template" in kra_map:
#         st.markdown("**Sample KRA Excel File**")
#         st.download_button(
#             label="📥 Sample KRA File",
#             data=kra_map["Default KRA Template"],
#             file_name="KRA.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )


# selected_template = st.selectbox("Choose a Template", template_options)
selected_template = template_options[0]
template_input = st.text_area("Email HTML Template", value=template_map[selected_template], height=300)

os.makedirs(".temp/kra_files/pdf_files", exist_ok=True)
os.makedirs(".temp/kra_files/excel_files", exist_ok=True)


st.subheader("📤 Upload Files to Attach to the email")
# Custom helper text
# st.info("👉 File name format: **Firstname Lastname.pdf**")
uploaded_pdfs = st.file_uploader(
    "Upload KRA PDF files ",
    type=["pdf", "xlsx"],
    accept_multiple_files=True,
)


if uploaded_pdfs:
    skipped_files = []
    added_files = []

    for file in uploaded_pdfs:
        ext = file.name.split(".")[-1].lower()

        # Route to correct folder
        if ext == "pdf":
            save_dir = ".temp/kra_files/pdf_files"
        elif ext == "xlsx":
            save_dir = ".temp/kra_files/excel_files"
        else:
            st.error(f"❌ Unsupported file type: {file.name}")
            continue

        save_path = os.path.join(save_dir, file.name)

        if os.path.exists(save_path):
            skipped_files.append(file.name)
        else:
            with open(save_path, "wb") as f:
                f.write(file.read())
            added_files.append(file.name)

    # Feedback to user
    if skipped_files:
        st.warning(f"⚠️ Skipped duplicates: {', '.join(skipped_files)}")
    if added_files:
        st.success(f"✅ Uploaded: {', '.join(added_files)}")

st.subheader("📤 Upload Excel File to Trigger Emails")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
# st.info("👉 File name format: **KRA.xlsx**")

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



# st.subheader("📂 Current KRA Files in '.temp/kra_files/'")

# pdf_files = sorted(glob.glob(".temp/kra_files/*.pdf"))

# if pdf_files:
#     for pdf_file in pdf_files:
#         filename = os.path.basename(pdf_file)
#         file_stats = os.stat(pdf_file)
#         file_size_kb = round(file_stats.st_size / 1024, 2)
#         file_modified = datetime.fromtimestamp(file_stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

#         with st.expander(f"📄 {filename}"):
#             st.markdown(f"**Size:** {file_size_kb} KB")
#             st.markdown(f"**Last Modified:** {file_modified}")

#             # Download button
#             with open(pdf_file, "rb") as f:
#                 st.download_button(
#                     label="⬇️ Download PDF",
#                     data=f,
#                     file_name=filename,
#                     mime="application/pdf"
#                 )

#             # Optional PDF preview
#             st.markdown("**Preview:**")
#             st.pdf(pdf_file) if hasattr(st, "pdf") else st.info("PDF preview not supported in this Streamlit version.")

#             # Delete button
#             if st.button(f"🗑️ Delete {filename}", key=f"delete_{filename}"):
#                 os.remove(pdf_file)
#                 st.success(f"Deleted {filename}")
#                 st.experimental_rerun()
# else:
#     st.info("No KRA PDF files found.")
