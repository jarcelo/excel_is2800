import sys
import re
import gspread
import time
from gspread_formatting import *
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('is2800-1166f40193c8.json', scope)

# Authorize
gsheet = gspread.authorize(credentials)
# Open workbook
workbook = gsheet.open('test')
# Get worksheets
worksheets = workbook.worksheets()   # Can fail starting here!//CHeck for optimization
  
#body = {}
body = {"requests": []}
def addTargetCell(sID, target):
    body['requests'].append(
        {"repeatCell": {
            "range": {
                "sheetId": sID,
                "startRowIndex": target - 1,
                "endRowIndex": target,
                "startColumnIndex": 1,
                "endColumnIndex": 4
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red" : 1,
                        "green": 0.9,
                        "blue": 0.9
                    }
                }
            },
            "fields": "userEnteredFormat(textFormat, backgroundColor)"  #remove textFormat?
        }})

#print(sheetIDList)

#Get the target list. Zero values from column D

zeroValues = []
#Loop through column D values and get the index for zero values
start = 12  
for index, item in enumerate(worksheets):
    source = str(item)
    #print("Source : " + source)
    if index >= start:
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        for index, item in enumerate(columnDItems):
            if item == '0'and index > 1 and index < (colDLength - 1):
                #Get the SID
                sID = re.search('id:(.+?)>', source).group(1)
                #zeroValues.append("D" + str(index + 1))
                #Get the target cell in Column D
                colDTarget = index + 1
                addTargetCell(sID, colDTarget)
                #print("ID: " + sID + "| Col : " + str(colDTarget))

#print(zeroValues)
print(body)

workbook.batch_update(body=body)