import streamlit as st
import pandas as pd
import smtplib
import time
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(
    page_title="Gmail HR Cold Mailer",
    layout="centered"
)

st.title("📧 Gmail HR Cold Email Sender")
st.caption("Gmail App Password • Smart Excel • Live Status • Safe Pause")

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------

def normalize(col):
    """Normalize column names for matching"""
    return re.sub(r"[\s_\-]", "", col.lower())


def detect_columns(df):
    """
    Auto-detect name, email, company columns
    Accepts messy headers like:
    HR_name, Hr Name, NAME, Email ID, Company Name, Organization, etc.
    """
    normalized = {normalize(c): c for c in df.columns}

    name_keys = ["name", "hrname", "candidate", "person"]
    email_keys = ["email", "mail"]
    company_keys = ["company", "organization", "org", "employer"]

    detected = {}

    for key, original in normalized.items():
        if "name" not in detected and any(k in key for k in name_keys):
            detected["name"] = original
        if "email" not in detected and any(k in key for k in email_keys):
            detected["email"] = original
        if "company" not in detected and any(k in key for k in company_keys):
            detected["company"] = original

    return detected

# --------------------------------------------------
# UI INPUTS
# --------------------------------------------------

excel_file = st.file_uploader(
    "📂 Upload Excel (any format)",
    type=["xlsx"]
)

resume_file = st.file_uploader(
    "📎 Upload Resume (PDF)",
    type=["pdf"]
)

sender_email = st.text_input("📧 Gmail Address")
sender_password = st.text_input(
    "🔐 Gmail App Password (NOT normal password)",
    type="password"
)

subject = st.text_input("✉️ Email Subject")

message_template = st.text_area(
    "📝 Email Body (use {name} & {company})",
    height=260
)

st.divider()

batch_size = st.selectbox("📦 Emails per batch", [50, 75, 100])
email_delay = st.slider("⏳ Delay between emails (seconds)", 2, 10, 4)
batch_delay = st.slider("⏸ Delay between batches (minutes)", 1, 20, 5)

# --------------------------------------------------
# SEND EMAILS
# --------------------------------------------------

if st.button("🚀 Start Sending Emails"):

    if not all([
        excel_file,
        sender_email,
        sender_password,
        subject,
        message_template
    ]):
        st.error("⚠️ Please fill all required fields")
        st.stop()

    # Read Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception:
        st.error("❌ Unable to read Excel file")
        st.stop()

    # Detect required columns
    detected = detect_columns(df)

    if not all(k in detected for k in ["name", "email", "company"]):
        st.error("❌ Could not auto-detect name, email, company columns")
        st.write("Detected columns:", list(df.columns))
        st.stop()

    # Standardize dataframe
    df = df.rename(columns={
        detected["name"]: "name",
        detected["email"]: "email",
        detected["company"]: "company"
    })

    df = df[["name", "email", "company"]].dropna()
    total = len(df)

    if total == 0:
        st.error("❌ No valid email records found")
        st.stop()

    # Read resume ONCE (mobile-safe attachment)
    resume_bytes = None
    if resume_file is not None:
        resume_bytes = resume_file.read()

    st.success("✅ Excel validated successfully")
    st.info(f"📊 Total Emails: {total}")

    # Live UI elements
    status_text = st.empty()
    progress_bar = st.progress(0)

    sent = 0

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        for i in range(total):
            row = df.iloc[i]

            # LIVE STATUS
            status_text.info(f"📨 Sending email {sent + 1} / {total}")

            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = row["email"]
            msg["Subject"] = subject

            body = message_template.format(
                name=row["name"],
                company=row["company"]
            )
            msg.attach(MIMEText(body, "plain"))

            # Attach resume if provided (mobile-safe)
            if resume_bytes is not None:
                attachment = MIMEBase("application", "pdf")
                attachment.set_payload(resume_bytes)
                encoders.encode_base64(attachment)
                attachment.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{resume_file.name}"'
                )
                msg.attach(attachment)

            try:
                server.sendmail(
                    sender_email,
                    row["email"],
                    msg.as_string()
                )

                sent += 1
                progress_bar.progress(sent / total)
                time.sleep(email_delay)

            except smtplib.SMTPException:
                # Graceful pause on Gmail limit or block
                status_text.warning(
                    f"⏸ Sending paused\n\n"
                    f"✅ {sent} out of {total} emails sent successfully"
                )
                server.quit()
                st.stop()

            # Batch delay handling
            if batch_size and sent % batch_size == 0 and sent < total:
                status_text.info(
                    f"⏸ Batch complete. Waiting {batch_delay} minutes..."
                )
                time.sleep(batch_delay * 60)

        server.quit()
        status_text.success(
            f"🎉 Completed successfully: {sent} / {total} emails sent"
        )

    except Exception as e:
        st.error(f"❌ Unexpected error: {e}")
