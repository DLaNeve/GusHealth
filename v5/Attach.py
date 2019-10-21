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

# READ CONFIG FILE
with open("config.json","r") as config_file:
    cfg = json.load(config_file)

smtp_server = cfg['smtpserver']
smtp_port = cfg['smtpport']
smtp_login = cfg['smtplogon']
smtp_password = cfg['smtppwd']
from_addr = cfg['from']
to_addr = cfg['recipient1']
cc_addr = cfg['recipient2']
group = [to_addr, cc_addr]
body = cfg['body2']

msg = MIMEMultipart()
msg['From'] = cfg['from']
msg['To'] = to_addr + ',' + cc_addr  # is this needed since we use the group
msg['Subject'] = cfg['subject2']
msg.attach(MIMEText(body, 'plain'))

filename = cfg['query_file']
attachment = open(filename, "rb")
p = MIMEBase('application', 'octet-stream')
p.set_payload((attachment).read())
encoders.encode_base64(p)
p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(p)
text = msg.as_string()

try:  # INITIALIZE AND LOGON TO SMTP SERVER
    smtp_server = smtplib.SMTP(smtp_server, smtp_port)
    smtp_server.ehlo()  # Send mandatory 'hello' msg
    smtp_server.starttls()  # Start TLS Encryption as we're not using SSL.
    smtp_server.login(smtp_login, smtp_password)  # login
    smtp_server.sendmail(from_addr, group, text)  # SEND THE EMAIL
    smtp_server.quit()
    print("Email(s) sent")
    print("Email logout")
except Exception:
    print("\n*** Email FAILED ***")
