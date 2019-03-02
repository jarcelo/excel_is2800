def getScaledScoreRequestBody(sID, target):
    body = {"requests": []}
    colDLength = len(target)
    target = colDLength + 1
    score = "=D" + str(target-1) + "/100*25"
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
    return body

##zeroValues = []
#Loop through column D values and get the index for zero values
#start = 0  
#for index, item in enumerate(worksheets):
#    source = str(item)
#    if index >= start:
#        columnDItems = item.col_values(4)
#        colDLength = len(columnDItems)
#        sID = re.search('id:(.+?)>', source).group(1)
#        colDTarget = colDLength + 1
#        getScaledScoreRequestBody(sID, colDTarget)

#print(body)
#workbook.batch_update(body=body)