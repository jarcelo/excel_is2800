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


body = {}
#def addTargetCell(sID, target):
body = (
    {"requests": [
        {"repeatCell": {
            "range": {
                "sheetId": 383978068,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 1,
                "endColumnIndex": 4
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold": True
                    },
                    "backgroundColor": {
                        "red" : 1,
                        "green": 0.9,
                        "blue": 0.9
                    }
                }
            },
            "fields": "userEnteredFormat(textFormat, backgroundColor)"
        }}
    ]}
)

body['requests'].append(
    {"repeatCell": {
            "range": {
                "sheetId": 1855958283,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 1,
                "endColumnIndex": 4
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold": True
                    },
                    "backgroundColor": {
                        "red" : 1,
                        "green": 0.9,
                        "blue": 0.9
                    }
                }
            },
            "fields": "userEnteredFormat(textFormat, backgroundColor)"
        }})

print(body)

workbook.batch_update(body=body)