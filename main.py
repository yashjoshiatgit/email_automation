import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# Configuration
SMTP_SERVER = "smtp.gmail.com"  # Use your SMTP server, e.g., "smtp.gmail.com" for Gmail
SMTP_PORT = 587
SENDER_EMAIL = "Your_Email"  # Replace with your email
PASSWORD = "Your_PassWord"  # Replace with your app password or email password (secure it)

# Load contacts from Excel file
def load_contacts(file_path):
    try:
        contacts = pd.read_excel(file_path)
        return contacts
    except Exception as e:
        print(f"Error reading contacts file: {e}")
        return None

# Load email template
def load_template(file_path):
    try:
        with open(file_path, 'r') as template_file:
            template = template_file.read()
        return template
    except Exception as e:
        print(f"Error reading template file: {e}")
        return None

# Customize the template with contact details
def create_message(template, subject, **kwargs):
    message = template
    for key, value in kwargs.items():
        message = message.replace(f"{{{{ {key} }}}}", value)
    return subject, message

# Send email
def send_email(receiver_email, subject, message):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure the connection
            server.login(SENDER_EMAIL, PASSWORD)
            server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())
        print(f"Email sent successfully to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")

# Main function to execute the automation
def main():
    contacts_file = Path("data/contacts.xlsx")
    template_file = Path("templates/email_template.txt")
    
    # Load data
    contacts = load_contacts(contacts_file)
    template = load_template(template_file)
    if contacts is None or template is None:
        print("Error loading files.")
        return

    # Process each contact
    for _, contact in contacts.iterrows():
        email = contact['Email']
        company_name = contact['Company Name']
        contact_person = contact.get('Contact Person', 'Hiring Team')
        location = contact['Location']

        # Customize the email
        subject, message = create_message(
            template,
            subject=f"Internship Inquiry at {company_name}",
            company_name=company_name,
            contact_person=contact_person,
            location=location
        )
        
        # Send the email
        send_email(email, subject, message)

if __name__ == "__main__":
    main()
