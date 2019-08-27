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
#workbook = gsheet.open('Excel_Chapter4_Scorecard_007')
workbook = gsheet.open('test')
# Get worksheets
worksheets = workbook.worksheets()   

body = {"requests": []}

def getScaledScore(sID, target):
    colDLength = len(target)
    target = colDLength + 1
    #score = "=D" + str(target-1) + "/100*25"
    score = "=D" + str(colDLength) + "/100*25"
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
                        "formulaValue": score
                    },
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True
                        }
                    }
                }]
            }],
            "fields": "*"
        }}
    )
    #return body

##zeroValues = []
#Loop through column D values and get the index for zero values
start = 11  
for index, item in enumerate(worksheets):
    source = str(item)
    if index >= start:
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        sID = re.search('id:(.+?)>', source).group(1)
        colDTarget = colDLength + 1
        getScaledScore(sID, columnDItems)

print(body)
workbook.batch_update(body=body)