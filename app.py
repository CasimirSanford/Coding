import gradio as gr
import smtplib
import ssl
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Function to read recipient list and return column options
def get_columns(file):
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file.name)
        elif file.name.endswith(".xls") or file.name.endswith(".xlsx"):
            df = pd.read_excel(file.name)
        else:
            return [], "❌ Unsupported file format! Please upload CSV, XLS, or XLSX."
        
        return list(df.columns), ""
    except Exception as e:
        return [], f"❌ Error reading file: {e}"

# Function to send emails
def send_bulk_emails(smtp_server, smtp_port, email_address, email_password, subject, message, file, email_column, attachment):
    # Read recipient list
    if not email_column:
        return "❌ Please select an email column!"
    
    columns, error = get_columns(file)
    if error:
        return error
    
    if email_column not in columns:
        return "❌ Selected column not found in file!"
    
    recipients = pd.read_csv(file.name) if file.name.endswith(".csv") else pd.read_excel(file.name)
    success_count = 0
    failed_emails = []

    for _, row in recipients.iterrows():
        to_email = row[email_column]
        name = row.get("name", "User")

        if pd.isna(to_email):
            continue

        try:
            msg = MIMEMultipart()
            msg["From"] = email_address
            msg["To"] = to_email
            msg["Subject"] = subject

            formatted_message = message.replace("\n", "<br>")
            html_content = f"""
            <html>
            <body style='text-align: center; font-family: Arial, sans-serif;'>
                <h2><b>Hello {name},</b></h2>
                <p style='font-size: 16px;'>{formatted_message}</p>
                <p style='color: blue;'><i>Best Regards,<br>Your Name</i></p>
            </body>
            </html>
            """
            msg.attach(MIMEText(html_content, "html"))

            # Attach file if provided
            if attachment:
                file_path = attachment.name
                with open(file_path, "rb") as file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
                    msg.attach(part)

            # Connect to SMTP server
            context = ssl.create_default_context()
            with smtplib.SMTP(smtp_server, int(smtp_port)) as server:
                server.starttls(context=context)
                server.login(email_address, email_password)
                server.sendmail(email_address, to_email, msg.as_string())

            success_count += 1
        except Exception as e:
            failed_emails.append(f"{to_email}: {e}")

    return f"✅ {success_count} emails sent!\n❌ Failed: {len(failed_emails)}\n" + "\n".join(failed_emails)


# Gradio Interface
def update_email_column(file):
    columns, error = get_columns(file)
    if error:
        return gr.update(choices=[], value=None), error
    return gr.update(choices=columns, value=None), ""

with gr.Blocks() as app:
    gr.Markdown("# 📩 Bulk Email Sender with Column Selection")
    
    smtp_server = gr.Textbox(label="SMTP Server", placeholder="smtp.gmail.com")
    smtp_port = gr.Textbox(label="SMTP Port", placeholder="587")
    email_address = gr.Textbox(label="Your Email Address")
    email_password = gr.Textbox(label="Your Email Password", type="password")
    subject = gr.Textbox(label="Email Subject")
    message = gr.Textbox(label="Email Message", lines=10)
    file = gr.File(label="Upload CSV/XLS/XLSX (name, email)")
    email_column = gr.Dropdown(label="Select Email Column", choices=[])
    attachment = gr.File(label="Attach File (Optional)", optional=True)
    
    file.change(update_email_column, inputs=file, outputs=[email_column])
    
    send_button = gr.Button("Send Emails")
    output_text = gr.Textbox(label="Output", interactive=False)
    
    send_button.click(
        send_bulk_emails, 
        inputs=[smtp_server, smtp_port, email_address, email_password, subject, message, file, email_column, attachment], 
        outputs=output_text
    )

app.launch()
