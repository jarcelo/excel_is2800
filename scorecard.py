import sys
import re
import gspread
import time
from gspread_formatting import *
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials
from remove_group_score import getGroupScoreRequestBody
from get_raw_score import getRawScoreRequestBody
from get_scaled_score import getScaledScoreRequestBody

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('is2800-1166f40193c8.json', scope)

# Authorize
gsheet = gspread.authorize(credentials)
# Open workbook
workbook = gsheet.open('test')
# Get worksheets
worksheets = workbook.worksheets()  

groupScoreBody = {}
rawScoreBody = ''
scaledScoreBody = ''
#Each function call will return a body?
start = 0  
for index, item in enumerate(worksheets):
    source = str(item)
    if index >= start:
        columnAItems = item.col_values(1)
        columnDItems = item.col_values(4)
        colDLength = len(columnDItems)
        sID = re.search('id:(.+?)>', source).group(1)
        colDTarget = colDLength + 1

        #function call
        groupScoreBody = getGroupScoreRequestBody(sID, columnAItems, columnDItems)
        #rawScoreBody = getRawScoreRequestBody(sID, columnDItems)
        #scaledScoreBody = getScaledScoreRequestBody(sID, columnDItems)

print(groupScoreBody)
#workbook.batch_update(body=groupScoreBody)
#workbook.batch_update(body=rawScoreBody)
#workbook.batch_update(body=scaledScoreBody)