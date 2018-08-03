import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import user_details as ud


def get_email_body(file_to_read):
    '''get the body of the email'''
    file = open(file_to_read, "r", encoding='utf-8-sig')
    body = file.read()
    file.close()
    return body


def send_email(toaddr, body, subject):
    '''(str, str)
    Given an address, the body, and the subject of the email, function sends
    an email automatically'''
    
    msg = MIMEMultipart()
    msg['From'] = ud.fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject
    # Convert the file to text
    # Removes extra characters from beginning of the list
    msg.attach(MIMEText(body, 'plain'))
    server = smtplib.SMTP(ud.host, ud.port)
    server.starttls()
    server.login(ud.fromaddr, ud.password)
    text = msg.as_string()
    server.sendmail(ud.fromaddr, toaddr, text)
    server.quit()
