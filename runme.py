import time

fileList = ['remove_group_score.py','identify_incorrect_items.py','get_raw_score.py', 'get_scaled_score.py']


for file in fileList:
    success = False
    while success == False:
        try:
            print("Executing script: " + file)
            exec(compile(open(file).read(), file, 'exec'))
            success = True
        except:
            print('Task cannot be completed. Will retry after 3 minutes...')
            time.sleep(180)