import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('is2800-1166f40193c8.json', scope)

# Authorize
gsheet = gspread.authorize(credentials)
# Open workbook
workbook = gsheet.open('test')
#workbook = gsheet.open('Excel_Case2_Scorecard_006')
# Get worksheets count
sheetCount = workbook.worksheets()

# Format cells with zero grade values; change the background color
bgfmt = CellFormat(
    backgroundColor=color(1, 0.9, 0.9)
)
scorefmt = CellFormat(
    textFormat = TextFormat(bold=True)
)

def removeItemGroupScore(columnAItems, columnDItems):
    print("Removing group score ... ")
    colDLength = len(columnAItems)
    for index, item in enumerate(columnAItems):
        if index > 1 and index < (colDLength - 1):
            targetCell = 'D' + str(index + 1)
            nextEarnedValue = columnDItems[index + 1]
            if item != "" and nextEarnedValue == '0':
                print("Removing " + targetCell + " value ...")
                worksheet.update_acell(targetCell, '')
    print("Done.")


def identifyIncorrectItem(items):
    print("Indentifying incorrect items ... ")
    colDLength = len(items)
    for index, item in enumerate(items):
        if item == '0'and index > 1 and index < (colDLength - 1) :
            cellRange = "B" + str(index + 1) + ":" + "D" + str(index + 1)
            format_cell_range(worksheet, cellRange, bgfmt)
    print("Done.")


def calculateScore(items):
    print("Calculating submission score ... ")
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


# Update scorecard
start = -1   # Replace variable with the last success index if API quota is reached before execution is finished
for index, item in enumerate(sheetCount):
    if index > start:
        print("Working on : "+ str(index) + "  " + str(item))
        worksheet = workbook.get_worksheet(index)
        columnAItems = worksheet.col_values(1)
        columnDItems = worksheet.col_values(4)

        #removeItemGroupScore(columnAItems, columnDItems)
        #identifyIncorrectItem(columnDItems)
        #calculateScore(columnDItems)

        print("----------------------------------------------------")