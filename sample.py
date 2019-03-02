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

requests.update(
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
    }}
)


requests.update(
    {"requests": [
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
        }}
    ]}
)




requests = {"requests": [
    {"repeatCell": {
        "range": {
            "sheetId": 1855958283,
            "startRowIndex": 0,
            "endRowIndex": 1
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "bold": True
                }
            }
        },
        "fields": "userEnteredFormat.textFormat.bold"
    }},
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



sheetIDList = []
#Get the SheetIDs
for sheets in worksheets:
    source = str(sheets)
    #print(source)
    try:
        sID = re.search('id:(.+?)>', source).group(1)
        sheetIDList.append(sID)
    except AttributeError:
        print ("Found nothing.")



        

'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "stringValue": "Joe"
                    }
                }]
            }],


'range': {
                "sheetId": sID,
                "startRowIndex": target,
                "endRowIndex": target + 1,
                "startColumnIndex": 3,
                "endColumnIndex": 4
            },
            'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "formulaValue": score
                    }
                }]
            }],


'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "formulaValue": score
                    },
                    "effectiveFormat": {
                        "borders": {
                            "top": {
                                "style": "SOLID_THICK"
                            }
                        }
                    }
                }]
            }],


 'rows': [{
                                "values": [{
                                    "userEnteredValue": {
                                        "stringValue": 'Joe'
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