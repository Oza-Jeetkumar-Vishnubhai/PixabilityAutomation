import smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from responseReading import read
from dotenv import load_dotenv 
import os

load_dotenv()
password=os.getenv('password')
sender_email=os.getenv('sender_email')

def send():
    email = read()[2]
    subject = "An email with attachment from Python"
    body = "Pixability Deck"
    sender_email = "oza.jeetkumar@miqdigital.com"
    receiver_email = email
    # password = "gfls ivkl qdjd mzjq"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  

    message.attach(MIMEText(body, "plain"))

    filename = "test_pixability.pptx"  

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    message.attach(part)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
    print("Deck sent successfully ✅")
