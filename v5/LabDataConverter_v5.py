#!/usr/bin/python3

import time
import json
import my_functions_v5 as mf

with open("config.json","r") as config_file:
    cfg = json.load(config_file)
    freq = cfg["mail_chk_freq"]

while True:
    mail = mf.login_mail()                              # login to GusHealth Email
    return_values = mf.check_for_lab_email(mail)        # chk for proper email-save the attachment
    success = return_values[0]                          # initialize
    filename = return_values[1]                         # initialize
 #   filename = "Gus ALM 2018-10.xlsx"               # testing only
 #   filename = "Tracy for Maria May 2018.xlsx"      # testing only
 #   success = 1                                     # testing only
    if success:
        query_date = mf.insert_data(filename)
        mf.create_csv_file(query_date)
        mf.send_mail_attachment()
        mf.cleanup()
    else:
        print ('Sleeping', freq, "seconds")
        time.sleep(int(freq))
