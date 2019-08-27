def getGroupScoreRequestBody(sID, columnAItems, columnDItems):
    colDLength = len(columnDItems)
    body = {"requests": []}
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
    return body

#zeroValues = []
#Loop through column D values and get the index for zero values
#start = 0  
#for index, item in enumerate(worksheets):
#    source = str(item)
#    if index >= start:
#        columnAItems = item.col_values(1)
#        columnDItems = item.col_values(4)
#        colDLength = len(columnDItems)
#        sID = re.search('id:(.+?)>', source).group(1)
#        colDTarget = colDLength + 1
#        addTargetCell(sID, columnAItems, columnDItems)

#print(body)
#workbook.batch_update(body=body)
