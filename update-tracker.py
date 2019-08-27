import sys
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
#workbook = gsheet.open('test')
workbook = gsheet.open('Excel_Case4_Scorecard_006')
# Get worksheets count
#sheetCount = workbook.worksheets()   # Can fail starting here!//CHeck for optimization
def getSheetCount(workbook):
    try:
        return workbook.worksheets()
    except:
        print("An error has occurred. API sheetcount read quota has been exhausted.")
        return 0

sheetCount = getSheetCount(workbook)
while sheetCount == 0:
    print('\nWill retry after 5 mins ...\n')
    time.sleep(300)
    sheetCount = getSheetCount(workbook)

# Format cells with zero grade values; change the background color
bgfmt = CellFormat(
    backgroundColor = Color(1, 0.9, 0.9)
)
scorefmt = CellFormat(
    textFormat = TextFormat(bold=True)
)


def removeItemGroupScore(columnAItems, columnDItems):
    print("Checking group score.")
    try:
        colDLength = len(columnAItems)
        target_list = []
        for index, item in enumerate(columnAItems):
            if index > 1 and index < (colDLength - 1):
                targetCell = 'D' + str(index + 1)
                nextEarnedValue = columnDItems[index + 1]
                if item != "" and nextEarnedValue == '0':
                    print("Removing " + targetCell + " value ...")
                    target_list.append(worksheet.acell(targetCell))
        for cell in target_list:
            cell.value = ''
        if len(target_list) > 0:
            worksheet.update_cells(target_list)
        print("Done.")

    except:
        print("\nAn error has occurred. Quota has been exhausted while removing group score.") 
        return 0
    

def identifyIncorrectItem(items):
    print("Indentifying incorrect items ... ")
    try:
        colDLength = len(items)
        for index, item in enumerate(items):
            if item == '0'and index > 1 and index < (colDLength - 1) :
                cellRange = "B" + str(index + 1) + ":" + "D" + str(index + 1)
                format_cell_range(worksheet, cellRange, bgfmt)
                format_
        
        format
        print("Done.")
    except:
        print("\nAn error has occurred. Quota has been exhausted while marking incorrect items.") 
        return 0


def calculateScore(items):
    print("Calculating submission score ... ")
    try:
        colDLength = len(items)
        percentTotal = '=sum(D2:D' + str(colDLength) + ')'
        scale = '=D' + str(colDLength + 1) + "/100*25"
        percentCell = 'D' + str(colDLength + 1)
        scoreCell = 'D' + str(colDLength + 2)
        scoreCellRange = scoreCell + ":" + scoreCell
        format_cell_range(worksheet, scoreCellRange, scorefmt)

        worksheet.update_acell(percentCell, percentTotal)
        worksheet.update_acell(scoreCell, scale)
        print("Done.")
    except:
        print("\nAn error has occurred. Quota has been exhausted while calculating score.") 
        return 0


# Update excel scorecard
# If execution does not finish, replace 'start' variable with the last success index.
start = 38  
for index, item in enumerate(sheetCount):
    if index >= start:
        print("Working on : "+ str(index) + "  " + str(item))
        worksheet = workbook.get_worksheet(index)   #Possible resource exhaustion point
        if len(worksheet.row_values(1)) == 0:
            print("Student does not have a submission ...")
            noSubmissionMsg = "No submission as of " + str(date.today())
            worksheet.update_acell('B1', noSubmissionMsg)
        else:
            columnAItems = worksheet.col_values(1)
            columnDItems = worksheet.col_values(4)

            #removeItemGroupScore(columnAItems, columnDItems)
            #identifyIncorrectItem(columnDItems)
            calculateScore(columnDItems)

            #itemScoreStatus = removeItemGroupScore(columnAItems, columnDItems)
            #while itemScoreStatus == 0:
            #    print('Will retry after 5 mins (' + str(time.ctime()) + ') ...\n')
            #    time.sleep(300) #Sleep for 5 mins
            #    itemScoreStatus = removeItemGroupScore(columnAItems, columnDItems)

            #incorrectItemStatus = identifyIncorrectItem(columnDItems)
            #while incorrectItemStatus == 0:
            #    print('Will retry after 5 mins (' + str(time.ctime()) + ') ...\n')
            #    time.sleep(300) #Sleep for 5 mins
            #    incorrectItemStatus = identifyIncorrectItem(columnDItems)

            #calculateScoreStatus = calculateScore(columnDItems)
            #while calculateScoreStatus == 0:
            #    print('Will retry after 5 mins (' + str(time.ctime()) + ') ...\n')
            #    time.sleep(300) #Sleep for 5 mins
            #   calculateScoreStatus = calculateScore(columnDItems)

        
        print("----------------------------------------------------")

print ("\n********All Done*********\n")