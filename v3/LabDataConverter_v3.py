###  Lab Data Converter v3  ###

    # IMPORTS
import time
import email
import imaplib
import subprocess
import json
import os
import sys
from my_functions_v3 import cleanup
from my_functions_v3 import send_mail
from my_functions_v3 import logit
from my_functions_v3 import convert_data

                        # OPEN/LOAD/CLOSE CONFIG FILE
with open("config.json","r") as config_file:
        cfg = json.load(config_file)

cleanup()

try:
    while True:   
                        # LOGON TO EMAIL SERVER 
        try:
            mail = imaplib.IMAP4_SSL(cfg["imapserver"])
            mail.login(cfg["imaplogon"],cfg["imappwd"])
            mail.select('inbox')
            print("Email login")
        except:  
            print("Unable to login to mail server...aborting")
            sys.exit(1)  # replace with retry someday

                        # CHECK FOR SPECIFIC EMAIL (From and Unopened)        
        result, msgs = mail.search(None,
                                   'UNSEEN',
                                   'FROM', cfg["imapfrom"])    
        msgs = msgs[0].split()                        
        for emailid in msgs:
            result, data = mail.fetch(emailid, '(RFC822)')
            email_body = data[0][1]
            m = email.message_from_bytes(email_body)
            if m.get_content_maintype() != 'multipart':
                continue
            for part in m.walk():            
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                filename=part.get_filename()         
                if cfg['attach_keyword'] in filename and 'xlsx' in filename:
                    print("New EMAIL w/attachment Found")
                    sv_path = os.path.join(cfg["local_save_dir"], filename)
                    fp = open(sv_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    print('Spreadsheet retrieved and saved')
                    
                        # CONVERT SPREADSHEET
                    convert_data(sv_path)

                        # COPY FILE (rawdata.csv) TO GOOGLE DRIVE
                    sv_path = os.path.join(cfg["local_save_dir"], "Samples.csv")
                    try:
                        subprocess.call(['rclone','copy',sv_path,'googledrive:'],-1)
                        print("File (samples.csv) copied to Google Drive")
                    except(exception):
                        print("*** ERROR COPYING FILE TO GOOGLE DRIVE")
                        print(exception)

                        # SEND EMAIL/SMS
                    send_mail()
                    cleanup()
                    
        else:
            print("*** No qualifying mail ***")
            print("Email Logout")
            print("... waiting...",cfg["mail_chk_freq"],"secs ***\n")
                           
        mail.logout()
        time.sleep(int(cfg["mail_chk_freq"]))
        
except KeyboardInterrupt:   # STOP UPON RECEIVING A CTRL C
        print("Operator chose to stop process.        " + time.asctime())
        exit()
    
   
