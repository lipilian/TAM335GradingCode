# %%import package
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# %%
# Create a multipart message and set headers
def sendEmail(subject, body, sender_email, receiver_email, password, attachmentsPath):
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = attachmentsPath  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
# %%
subject = "Your Comments for Full report Block 1"
body = "This is email about the comment of your full report 1 (attachment). "  \
    "If you have any question, " + \
    "please contact by my university email: liuhong2@illinois.edu"
sender_email = "liuhong2xu@gmail.com"
receiver_email = "liuhong2@illinois.edu" # TODO: change email
attachmentsPath = './sheet/Hong_Liu_liuhong2.pdf'
password = 'Aa12233445!'
sendEmail(subject, body, sender_email, receiver_email, password, attachmentsPath)
# %%
