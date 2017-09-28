#----------------------------------
# TEAM 356 SCOUTING APP Version 1.0
# Author: Satvir Sandhu
#
# *** VEXDB is used to gather tournament stats and values, as matrix calculations
# don't need to be done locally ***
#
# The purpose of this application is to simplify the scouting process.
#
# Simplification is done by displaying OPR, DPR, CCWM, and regular stats
# onto a spreadsheet which can be manipulated in Office Excel.
#
# Prior to running the program, the user must open the file and change
# the "sku" value, to the code of the event that the team is competing at.
#
# The Excel spreadsheet included in the package must be located and opened.
#
# WINDOWS and Python 3 ONLY!
#----------------------------------
from openpyxl import *
from tkinter import filedialog
from urllib import request
from bs4 import BeautifulSoup
import urllib.parse
import json
import requests
import time


# Opens a Windows Explorer dialog to locate spreadsheet"
scoutingFile = load_workbook(filedialog.askopenfilename(), guess_types=True)
# Asks user which sheet to access
print(scoutingFile.get_sheet_names())
choice = input('What sheet would you like to access?')
chosenSheet = scoutingFile[choice]

# Accesses VEXDB and gathers all data regarding the specific event and prints
# to the spreadsheet
while True:

    sku = 'RE-VRC-15-3713'
    dest_URL = 'https://api.vexdb.io/v1/get_rankings'

    final_URL = (dest_URL +'?sku=' +sku)
    print(final_URL)
    data = requests.get(final_URL).json()
    print(data)
    c = 2
    for item in data['result']:
        print("---------------------------------------")
        print ('Rank: ' +item['rank'])
        if (int(item['rank']) <= 8):
            chosenSheet['B' + str(c)] = (item['rank'] + 'PP**')
        else:
            chosenSheet['B' + str(c)] = (item['rank'])
        print ('Team: ' +item['team'])
        chosenSheet['C' + str(c)] = item['team']
        print ("SP: " +item['sp'])
        chosenSheet['D' + str(c)] = item['sp']
        print("TRSP: " + str(item["trsp"]))
        chosenSheet['E' + str(c)] = item['trsp']
        print("Max Score: " +item['max_score'])
        chosenSheet['F' + str(c)] = item['max_score']
        print("OPR: " + str(item["opr"]))
        chosenSheet['G' + str(c)] = item['opr']
        print("DPR: " + str(item["dpr"]))
        chosenSheet['H' + str(c)] = item['dpr']
        print("CCWM: " + str(item["ccwm"]))
        chosenSheet['I' + str(c)] = item['ccwm']

        # Counter ensures that sheets aren't over-written.
        c += 1
        # File is saved.
        scoutingFile.save("Team356_Scouting.xlsx")
    time.sleep(30)
