import smtplib
import imaplib
import email
import os
import pandas as pd
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

  # OPEN/LOAD/CLOSE CONFIG FILE
with open("config.json","r") as config_file:
    cfg = json.load(config_file)

def cleanup():
    try:
        os.remove(cfg["local_save_dir"] + "/rawdata.xlsx") 
        os.remove(cfg["local_save_dir"] + "/Samples.csv")
        print("Cleanup completed\n")
    except:
        print("No cleanup necessary")
    
def login_mail(mail):
    try:
        mail = imaplib.IMAP4_SSL(cfg["imapserver"])
        mail.login(cfg["imaplogon"],cfg["imappwd"])
        mail.select('inbox')
        print("Email login")
    except:  
        print("Unable to login to mail server...aborting")
        sys.exit(1)  # replace with retry someday!                              

def get_attach(msgs,mail):    
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
                    


def convert_data(sv_path):
    os.chdir(cfg["local_save_dir"])
    os.rename(sv_path,"rawdata.xlsx")
    df = pd.read_excel ("rawdata.xlsx")
    df = df.drop('Totals', axis=1)
    df = df.dropna(how='all', axis=0)
    df.index.rename('Practice', inplace=True)
    df = df.dropna(subset=['DOCTOR'])
    df = df[df.index.notnull()]
    df = df.dropna(how='any', axis=1)
    df = pd.melt(df.reset_index(), id_vars=['Practice','DOCTOR'], var_name='Date', value_name='Samples')
    df['Samples'] = df['Samples'].astype(int)
    try:
        df.to_csv('Samples.csv')
        print("File converted & saved locally")
    except:
        print("*** File conversion error *** ")

def send_mail():   
    smtp_server =       cfg['smtpserver']
    smtp_port =         cfg['smtpport']
    smtp_login =        cfg['smtplogon']
    smtp_password =     cfg['smtppwd']
    to_addr =           cfg['recipient1']
    cc_addr =           cfg['recipient2']
    group =             [to_addr, cc_addr]
    from_addr =         cfg['from']
    subject_line =      cfg['subject']
    body_msg =          cfg['body1']  
    msg =               MIMEMultipart()
    msg['From'] =       from_addr
    msg['To'] =         to_addr + ',' + cc_addr
    msg['Subject'] =    subject_line 
    body =              body_msg
           
    msg.attach(MIMEText(body, 'plain'))
    text = msg.as_string()
    try: #  INITIALIZE AND LOGON TO SMTP SERVER
        smtp_server = smtplib.SMTP(smtp_server, smtp_port) # Specify Gmail Mail server
        smtp_server.ehlo()  # Send mandatory 'hello' message to SMTP server
        smtp_server.starttls() # Start TLS Encryption as we're not using SSL.
        smtp_server.login(smtp_login,smtp_password)# login
        smtp_server.sendmail(from_addr, group, text)   # SEND THE EMAIL
        smtp_server.quit()
        print("Email(s) sent")
        print("Email logout")
    except Exception:
        print("\n*** Email FAILED ***")


def logit(log_data):
    print(log_data[0],log_data[1],log_data[2],log_data[3]) # write to console
    loginfo = {
            "Date":     log_data[0],
            "Time":     log_data[1],
            "Event":    log_data[2],
            "Duration": log_data[3]
              }                  
    with open("/var/www/html/log.json","w+") as log_file:
        log_file.write("[")
        logrecord = ("\n" + json.dumps(loginfo) + "]")
        log_file.write(logrecord)    
            
    

##gmailaddress = 'reports@gushealth.com' #input("what is your gmail address? \n ")
##gmailpassword = 'GusHealth!' #input("what is the password for that email address? \n  ")
##mailto = '8138103281@vtext.com' #input("what email address do you want to send your message to? \n ")
##msg = 'Dashboards have been updated.' #input("What is your message? \n ")
##mailServer = smtplib.SMTP('smtp.gmail.com' , 587)
##mailServer.starttls()
##mailServer.login(gmailaddress , gmailpassword)
##mailServer.sendmail(gmailaddress, mailto , msg)
##print(" \n Sent!")
##mailServer.quit()
        
