import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('is2800-1166f40193c8.json', scope)

# Authorize
gsheet = gspread.authorize(credentials)
# Open workbook
workbook = gsheet.open('test')
# Get worksheets count
sheetCount = workbook.worksheets()
# Format cells with zero grade values; change the background color
bgfmt = CellFormat(
    backgroundColor=color(1, 0.9, 0.9)
)
scorefmt = CellFormat(
    textFormat = TextFormat(bold=True)
)

# Update grade tracker
for index, item in enumerate(sheetCount):
    print("Working on :"+ str(item))
    # Get the worksheet
    worksheet = workbook.get_worksheet(index)
    # Get the column D values
    items = worksheet.col_values(4)  
    # Get the column number for last entry in column D
    colDLength = len(items)
    print("Identifying incorrect items ... ")
    # Loop through the items. Change the background color if the values is zero
    for index, item in enumerate(items):
        if item == '0'and index != 1 and index != colDLength - 1 :
            cellRange = "B" + str(index + 1) + ":" + "D" + str(index + 1)
            format_cell_range(worksheet, cellRange, bgfmt)
            #print(index, item)  # Uncomment to print the items with zero score

    percentTotal = '=sum(D2:D' + str(colDLength) + ')'
    scale = '=D' + str(colDLength + 1) + "/100*25"
    percentCell = 'D' + str(colDLength + 1)
    scoreCell = 'D' + str(colDLength + 2)
    scoreCellRange = scoreCell + ":" + scoreCell
    format_cell_range(worksheet, scoreCellRange, scorefmt)

    #Get sum and score
    print("Calculating submission score ... ")
    worksheet.update_acell(percentCell, percentTotal)
    worksheet.update_acell(scoreCell, scale)
    print("Done.")
    print("------------------------------------------------")