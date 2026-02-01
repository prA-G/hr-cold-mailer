import streamlit as st
import pandas as pd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(
    page_title="Universal HR Cold Mailer",
    layout="centered"
)

st.title("📧 Universal HR Cold Email Sender")
st.caption("Reusable • Batch-based • Resume supported • Safe sending")

# -------------------- INPUT SECTION --------------------

excel_file = st.file_uploader(
    "📂 Upload Excel (name | company | email)",
    type=["xlsx"]
)

resume_file = st.file_uploader(
    "📎 Upload Resume (PDF only)",
    type=["pdf"]
)

sender_email = st.text_input("📧 Your Email ID")
sender_password = st.text_input(
    "🔐 Gmail App Password (not normal password)",
    type="password"
)

subject = st.text_input("✉️ Email Subject")

message_template = st.text_area(
    "📝 Email Body (use {name} and {company})",
    height=250
)

st.divider()

batch_size = st.selectbox(
    "📦 Emails per batch",
    [50, 75, 100]
)

email_delay = st.slider(
    "⏳ Delay between emails (seconds)",
    2, 10, 4
)

batch_delay = st.slider(
    "⏸ Delay between batches (minutes)",
    1, 20, 5
)

# -------------------- SEND BUTTON --------------------

if st.button("🚀 Start Sending Emails"):

    if not all([
        excel_file,
        resume_file,
        sender_email,
        sender_password,
        subject,
        message_template
    ]):
        st.error("⚠️ Please fill all fields")
        st.stop()

    try:
        df = pd.read_excel(excel_file)
    except Exception:
        st.error("❌ Invalid Excel file")
        st.stop()

    required_cols = {"name", "company", "email"}
    if not required_cols.issubset(df.columns):
        st.error("❌ Excel must contain: name, company, email")
        st.stop()

    total_emails = len(df)
    st.info(f"📊 Total HR Emails: {total_emails}")

    progress_bar = st.progress(0)
    sent_count = 0

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        for i in range(0, total_emails, batch_size):

            batch_df = df.iloc[i:i + batch_size]
            st.write(f"📦 Sending batch {i // batch_size + 1}")

            for _, row in batch_df.iterrows():

                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = row["email"]
                msg["Subject"] = subject

                body = message_template.format(
                    name=row["name"],
                    company=row["company"]
                )

                msg.attach(MIMEText(body, "plain"))

                resume = MIMEApplication(
                    resume_file.read(),
                    _subtype="pdf"
                )
                resume.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=resume_file.name
                )
                msg.attach(resume)

                server.sendmail(
                    sender_email,
                    row["email"],
                    msg.as_string()
                )

                sent_count += 1
                progress_bar.progress(sent_count / total_emails)

                time.sleep(email_delay)

            if i + batch_size < total_emails:
                st.warning(
                    f"⏸ Waiting {batch_delay} minutes before next batch"
                )
                time.sleep(batch_delay * 60)

        server.quit()
        st.success(f"✅ Successfully sent {sent_count} emails")

    except Exception as e:
        st.error(f"❌ Error: {e}")
