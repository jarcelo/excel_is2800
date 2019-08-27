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

def addTargetCell(sID, columnAItems, columnDItems):
    for index, item in enumerate(columnAItems):
            if index > 1 and index < (colDLength - 1):
                target = index + 1
                nextEarnedValue = columnDItems[index + 1]
                if item != "" and nextEarnedValue == '0':
                    body['requests'].append(
                        {"updateCells": {
                            'range': {
                                "sheetId": sID,
                                "startRowIndex": target - 1 ,
                                "endRowIndex": target,
                                "startColumnIndex": 3,
                                "endColumnIndex": 4
                            },
                            'rows': [{
                                "values": [{
                                    "userEnteredValue": {
                                        "stringValue": ''
                                    }
                                }]
                            }],
                            "fields": "*"
                        }}
                    )

start = 0  
for index, item in enumerate(worksheets):
    source = str(item)
    if index >= start:
        columnAItems = item.col_values(1)
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        sID = re.search('id:(.+?)>', source).group(1)
        colDTarget = colDLength + 1
        addTargetCell(sID, columnAItems, columnDItems)

print(body)
workbook.batch_update(body=body)
