###  Lab Data Converter v2  ###

    # IMPORTS
import time
import email
import imaplib
#import pandas as pd
#import subprocess
#import threading
import json
import os
import sys
from my_functions import send_mail
from my_functions import logit

                        # OPEN/LOAD/CLOSE CONFIG FILE
with open("config.json","r") as config_file:
        cfg = json.load(config_file)
#svdir = '/home/pi/Sean/v2'        # add this to config
try:
    while True:   
                        # LOGON TO EMAIL SERVER 
        try:
            mail = imaplib.IMAP4_SSL(cfg["imapserver"])
            mail.login(cfg["imaplogon"],cfg["imappwd"])
            mail.select('inbox')
            print("Email account logon successful")
        except:  
            print("Unable to logon to mail server...aborting")
            sys.exit(1)  # replace with retry

                        # CHECK FOR SPECIFIC EMAIL (From and Unopened)        
        result, msgs = mail.search(None,
           #                        'UNSEEN',
                                   'FROM', cfg["imapfrom"])
        
        msgs = msgs[0].split()                        
        for emailid in msgs:
            result, data = mail.fetch(emailid, '(RFC822)')
            email_body = data[0][1]
            m = email.message_from_bytes(email_body)
            if mail.is_multipart():
                print("multipart")
                for part in mail.walk():
                    if ctype == 'image/jpeg':
                        
                        filename=part.get_filename()
                        print(filename)
                
            if filename == "MARIA MARCH 2018.xlsx":
                    sv_path = os.path.join(cfg["local_save_dir"], filename)
                    fp = open(sv_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    print(sv_path,' SUCCESSFULLY SAVED')
                    
                        
            # test to determine if file exists locally 
#                extension = '.xlsx'
#                if os.path.splitext(filename)[-1] == cfg["extension"]:   
#                    if not os.path.isfile(sv_path):        
#                        
                        
#                    else:
#                        print('**** File already exists ****')

        else:
            print("No qualifying message(s). Will check again in:",cfg["mail_chk_freq"],"secs")            
                           
        mail.logout()
        time.sleep(int(cfg["mail_chk_freq"]))
        
except KeyboardInterrupt:   # STOP UPON RECEIVING A CTRL C
        print("Operator chose to stop process.        " + time.asctime())
        exit()
           
        
        #convert_data
        #copy to coud: subprocess.call(['rclone','copy','/home/pi/Sean/Samples.csv','googledrive:'],-1)
        #delete local saved spreadsheet
        #UPDATE THE LOG
        #log_data = [pf_date,pf_time,"Power Restored",mins]
        #logit(log_data)

        # SEND NOTIFICATION (email / SMS)            
        #emailThread=threading.Thread(target=send_mail)
        #emailThread.start()

    
   
