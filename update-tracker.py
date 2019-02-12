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

def identifyIncorrectItem(items):
    print("Indentifying incorrect items ... ")
    colDLength = len(items)
    for index, item in enumerate(items):
        if item == '0'and index > 1 and index < (colDLength - 1) :
            cellRange = "B" + str(index + 1) + ":" + "D" + str(index + 1)
            format_cell_range(worksheet, cellRange, bgfmt)
    print("Done.")


def calculateScore(colDLength):
    print("Calculating submission score ... ")
    percentTotal = '=sum(D2:D' + str(colDLength) + ')'
    scale = '=D' + str(colDLength + 1) + "/100*25"
    percentCell = 'D' + str(colDLength + 1)
    scoreCell = 'D' + str(colDLength + 2)
    scoreCellRange = scoreCell + ":" + scoreCell
    format_cell_range(worksheet, scoreCellRange, scorefmt)

    worksheet.update_acell(percentCell, percentTotal)
    worksheet.update_acell(scoreCell, scale)
    print("Done.")


# Update grade tracker
start = -1
for index, item in enumerate(sheetCount):
    if index > start:
        print("Working on : "+ str(index) + "  " + str(item))
        # Get the worksheet
        worksheet = workbook.get_worksheet(index)
        #worksheet = workbook.get_worksheet(i)
        # Get the column D values
        items = worksheet.col_values(4)
        #print("Indentifying incorrect items ... ")
        #identifyIncorrectItem(items)
        #for index, item in enumerate(items):
        #    if item == '0'and index > 1 and index < (colDLength - 1) :
        #        cellRange = "B" + str(index + 1) + ":" + "D" + str(index + 1)
        #        format_cell_range(worksheet, cellRange, bgfmt)
        #        for
        #        print(index, item)  # Un/comment to print the items with zero score

        #print(worksheet.col_values(4))
        # Get the column number for last entry in column D
        colDLength = len(items)
        calculateScore(colDLength)
     
        #print(worksheet.col_values(4))
        #i+=1
        #print("Current sheet id: " + str(1))
        print("----------------------------------------------------")