import sys
import re
import gspread
import time
from gspread_formatting import *
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials
from input_file import TargetWorkbook

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('is2800-1166f40193c8.json', scope)

# Authorize
gsheet = gspread.authorize(credentials)
# Open workbook
workbook = gsheet.open(TargetWorkbook)
# Get worksheets
worksheets = workbook.worksheets()   

body = {"requests": []}

def getRawScore(sID, target):
    colDLength = len(target)
    target = colDLength + 1
    percentage = '=sum(D2:D' + str(colDLength) + ')'
    body['requests'].append(
        {"updateCells": {
            'start': {
                "sheetId": sID,
                "rowIndex": target - 1,
                "columnIndex": 3
            },
            'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "formulaValue": percentage
                    }
                }]
            }],
            "fields": "*",
        }}
    )

start = 0  
for index, item in enumerate(worksheets):
    source = str(item)
    if index >= start:
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        sID = re.search('id:(.+?)>', source).group(1)
        colDTarget = colDLength + 1
        getRawScore(sID, columnDItems)

print(body)
workbook.batch_update(body=body)