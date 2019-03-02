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
worksheets = workbook.worksheets()

body = {"requests": []}

def addTargetCell(sID, target):
    percentage = "=sum(D2:D" + str(target-1) + ")"
    body['requests'].append(
        {"updateCells": {
            'range': {
                "sheetId": sID,
                "startRowIndex": target - 1,
                "endRowIndex": target,
                "startColumnIndex": 3,
                "endColumnIndex": 4
            },
            'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "formulaValue": percentage
                    },
                    "userEnteredFormat": {
                        "borders": {
                            "top": {
                                "style": "SOLID_MEDIUM"
                            }
                        }
                    }
                }]
            }],
            "fields": "*"
        }}
    )

zeroValues = []
#Loop through column D values and get the index for zero values
start = 12  
for index, item in enumerate(worksheets):
    source = str(item)
    if index >= start:
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        sID = re.search('id:(.+?)>', source).group(1)
        colDTarget = colDLength + 1
        addTargetCell(sID, colDTarget)

print(body)
workbook.batch_update(body=body)