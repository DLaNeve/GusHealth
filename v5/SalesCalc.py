#import os
import sqlite3
#import json

#with open ("config.json", "r") as config_file:
#    cfg = json.load (config_file)
def sales_calc():
    dbName = "GusHealth_v6.db"
    con = sqlite3.connect (dbName)
    c = con.cursor ()

    x = '2018-01-%'
    total_samples = 0

    for z in range (1, 10):
        rep_total = 0
        c.execute ('SELECT * '
                   'FROM SalesActivity '
                   'where date like (?) and '
                   'repno =(?)', (x, z))
        for row in c:
            rep_total = rep_total + row[6]
            print ([row], row[6] * row[7])
        print ('Rep Total =', rep_total)
        total_samples = total_samples + rep_total
    print ('Month Total =', total_samples)
