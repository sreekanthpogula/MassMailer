import smtplib
import logging
import os
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
from langchain_core.prompts import ChatPromptTemplate
from langchain.chains import RetrievalQA
from langchain_community.embeddings import OllamaEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_community.document_loaders import PyPDFDirectoryLoader
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_community.llms import ollama

# Load environment variables
load_dotenv()

config = {
    'MAIL_SERVER': os.getenv('MAIL_SERVER'),
    'MAIL_PORT': int(os.getenv('MAIL_PORT', 25)),
    'MAIL_USERNAME': os.getenv('MAIL_USERNAME'),
    # 'MAIL_PASSWORD': os.getenv('MAIL_PASSWORD'),
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
        kra_path = os.path.join("kra_files", kra_file_name)
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
        kra_path = os.path.join('kra_files', kra_file_name)

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
st.subheader("Email Template Selection")

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


# --- LLM Configuration with RetrievalQA ---
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

    # Vector Store Setup
    st.session_state.embeddings = OllamaEmbeddings(model="nomic-embed-text")
    st.session_state.loader = PyPDFDirectoryLoader("./data")
    st.session_state.docs = st.session_state.loader.load()
    st.session_state.text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
    st.session_state.final_documents = st.session_state.text_splitter.split_documents(st.session_state.docs[:1])
    st.session_state.vectors = FAISS.from_documents(
        st.session_state.final_documents,
        st.session_state.embeddings
    )

    # Create Retriever from FAISS
    st.session_state.retriever = st.session_state.vectors.as_retriever(search_kwargs={"k": 3})

    # Define LLM
    st.session_state.llm = ollama.Ollama(model="gemma3", temperature=0.7)

    # Define RetrievalQA chain
    st.session_state.qa_chain = RetrievalQA.from_chain_type(
        llm=st.session_state.llm,
        retriever=st.session_state.retriever,
        return_source_documents=True
    )

# --- Function to ask the assistant ---
def ask_bot(question):
    try:
        result = st.session_state.qa_chain.invoke({"query": question})
        answer = result["result"]
        # sources = result.get("source_documents", [])

        # # Format the sources nicely
        # if sources:
        #     source_texts = "\n\n".join(f"🔹 **Source {i+1}:**\n{doc.page_content.strip()}" for i, doc in enumerate(sources))
        #     answer += "\n\n---\n**Context from documents:**\n" + source_texts

        return answer
    except Exception as e:
        return f"Error: {e}"


# --- Sidebar UI ---
st.sidebar.title("Mass Mailer Assistant")
st.sidebar.markdown("Ask questions about the Mass Mailer application or get help with email templates.")

with st.sidebar.expander("💬 Ask the Assistant", expanded=False):
    user_input = st.chat_input("Ask something...")

if user_input:
    st.session_state.chat_history.append({"role": "user", "content": user_input})
    st.sidebar.chat_message("user").markdown(user_input)

    with st.spinner("Assistant is thinking..."):
        response = ask_bot(user_input)

    st.session_state.chat_history.append({"role": "assistant", "content": response})
    st.sidebar.chat_message("assistant").markdown(response)