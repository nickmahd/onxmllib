import smtplib
from email.message import EmailMessage

def email(message, to_email="venkat@oncept.net", subject='ME_watch Error', from_email="Me_watch@oncept.net", server='10.10.0.222'):
    msg = EmailMessage()
    msg['To'] = ', '.join(to_email)
    msg['From'] = from_email
    msg['Subject'] = subject
    msg.set_content(message)

    server = smtplib.SMTP(server)
    server.send_message(msg)
    server.quit()