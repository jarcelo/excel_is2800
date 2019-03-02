body = {"requests": []}

def getRawScoreRequestBody(sID, target):
    colDLength = len(target)
    target = colDLength + 1
    percentage = "=sum(D2:D" + str(target-1) + ")"
    body['requests'].append(
        {"updateCells": {
            'range': {
                "sheetId": sID,
                "startRowIndex": target - 1,
                "endRowIndex": target,
                "startColumnIndex": 3,
                "endColumnIndex": 4
            },
            'rows': [{
                "values": [{
                    "userEnteredValue": {
                        "formulaValue": percentage
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
        }}
    )
    return body

#zeroValues = []
#Loop through column D values and get the index for zero values
#start = 12  
#for index, item in enumerate(worksheets):
#    source = str(item)
#    if index >= start:
#        columnDItems = item.col_values(4)
#        colDLength = len(columnDItems)
#        sID = re.search('id:(.+?)>', source).group(1)
#        colDTarget = colDLength + 1
#        addTargetCell(sID, colDTarget)

#print(body)
#workbook.batch_update(body=body)