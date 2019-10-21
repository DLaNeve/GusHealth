import json
import openpyxl
import os
import sqlite3

with open ("config.json", "r") as config_file:
    cfg = json.load (config_file)

try:
    dbName = os.path.join (cfg["data_dir"], "GusHealth.db")
    con = sqlite3.connect (dbName)
    c = con.cursor ()
except:
    print ("ERROR:  Unable to connect to DB")

##Open the workbook   NEED TO DETERMINE WHEN AND WHAT THE WB NAME WILL BE  Also, what sheet!
wb = openpyxl.load_workbook (os.path.join (cfg["data_dir"], 'Gus ALM 2018-10.xlsx'), data_only=True)
ws = wb["April"]          # Need to determine what sheet
Month = ws.title

##Find the row containing the words "Referring... (block 1)
col_num = 1
for block1_start in range (1, 500):
    if ws.cell (block1_start, col_num).value == "Referring Provider Name":
        block1_start = block1_start + 1
        break

##Find the last row containing data in block 1.
col_num = 1
for block1_stop in range (block1_start, 500):
    if ws.cell (block1_stop, col_num).value is None:
        block1_stop = block1_stop - 1
        break

##Find the row containing the word "Referring... again  (block 2).
col_num = 1
for block2_start in range (block1_stop, 500):
    test = ws.cell (block2_start, col_num).value  # debug
    if ws.cell (block2_start, col_num).value == "Referring Provider Name":
        break

##Find the last row containing data in block 2
col_num = 1
for block2_stop in range (block2_start, 500):
    if ws.cell (block2_stop, col_num).value is None:
        block2_stop = block2_stop - 1
        break


for z in range (block1_start, block1_stop + 1):
    row_buffer = []
    for y in range (1, 10):  # col 1 thru col 9 of the first block of data
        row_buffer.append (ws.cell (z, y).value)
    c.execute ("""INSERT or REPLACE INTO selfpay VALUES(?,?,?,?)""",
               (Month, row_buffer[7], row_buffer[1], row_buffer[0]))

con.commit ()
con.close ()
