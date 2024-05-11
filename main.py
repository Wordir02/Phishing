import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587
smtp_username = 'dirigenzaitsumbria@outlook.it'
smtp_password = 'Wordir153275988!'
from_email = 'dirigenzaitsumbria@outlook.it'
to_email = 'allievo_Terlizzi.D@itsumbria.it'
subject = 'Report Percentuale Di presenza per {name}'
body = 'Per visualizzare il report delle presenze scarica il file excel allegato.'
print('Invio email...')

# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

# Add body to email
message.attach(MIMEText(body, "plain"))

# Open file in binary mode
filename = "test.ps1"  # In same directory as script
with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as pdf attachment
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()

# Log in to server using secure context and send email
with smtplib.SMTP(smtp_server, smtp_port) as smtp:
    smtp.starttls()
    smtp.login(smtp_username, smtp_password)
    smtp.sendmail(from_email, to_email, text)

