import os
import sqlite3
import email
import smtplib
import imaplib
import json
import openpyxl
import datetime
import time

from openpyxl.utils import get_column_letter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase


with open("config.json","r") as config_file:
    cfg = json.load(config_file)

attachfile=['all_samples_nonzero.csv']
for file in attachfile:
    try:
        with open(file,'rb') as fp:
            xxx=MIMEBase('application', 'octet-stream')
            xxx.set_payload(fp.read())
        encoders.encode_base64(xxx)
        xxx.add_header('Content-Disposition', 'attachment',filename=os.path.basename(file))
    except:
        print('unable to open file')
        raise

smtp_server = cfg['smtpserver']
smtp_port = cfg['smtpport']
smtp_login = cfg['smtplogon']
smtp_password = cfg['smtppwd']
to_addr = cfg['recipient1']
cc_addr = cfg['recipient2']
group = [to_addr, cc_addr]
from_addr = cfg['from']
subject_line = cfg['subject']
body_msg = cfg['body1']
msg = MIMEMultipart()
msg['From'] = from_addr
msg['To'] = to_addr + ',' + cc_addr
msg['Subject'] = subject_line
body = body_msg
msg.attach(xxx)

msg.attach(MIMEText(body, 'plain'))
text = msg.as_string()
try:  # INITIALIZE AND LOGON TO SMTP SERVER
    smtp_server = smtplib.SMTP(smtp_server, smtp_port)  # Specify Gmail Mail server
    smtp_server.ehlo()  # Send mandatory 'hello' message to SMTP server
    smtp_server.starttls()  # Start TLS Encryption as we're not using SSL.
    smtp_server.login(smtp_login, smtp_password)  # login
    smtp_server.sendmail(from_addr, group, text)  # SEND THE EMAIL
    smtp_server.quit()
    print("Email(s) sent")
    print("Email logout")
except Exception:
    print("\n*** Email FAILED ***")
