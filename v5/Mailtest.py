#!/usr/bin/env python
# encoding: utf-8

import os
import smtplib
import json
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

with open("config.json","r") as config_file:
    cfg = json.load(config_file)

COMMASPACE = ', '


def main():
    sender = cfg['smtplogon']
    gmail_password = cfg['smtppwd']
    recipients = cfg['recipient2']

    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = cfg['subject']
    outer['To'] = cfg['recipient2']
    outer['From'] = cfg['from']
    outer.preamble = 'You will not see this in a MIME-aware mail reader.\n'

    # List of attachments
    attachments = ['all_samples_nonzero.csv']

    # Add the attachments to the message
    for file in attachments:
        try:
            with open(file, 'rb') as fp:
                msg = MIMEBase('application', "octet-stream")
                msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            outer.attach(msg)
        except:
            print("Unable to open one of the attachments. Error: ", sys.exc_info()[0])
            raise

    composed = outer.as_string()

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, gmail_password)
            s.sendmail(sender, recipients, composed)
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ", sys.exc_info()[0])
        raise


if __name__ == '__main__':
    main()

