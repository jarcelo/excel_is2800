# This util will delete the files recursively within the folder

import os, shutil

#folder = "C:\\Users\\Master Joe\\Desktop\\Excel_Case_Submissions_006"
folder = "C:\\Users\\Master Joe\\Desktop\\test"

for the_file in os.listdir(folder):
    file_path = os.path.join(folder, the_file)
    try:
        if os.path.isfile(file_path):
            os.unlink(file_path)
        elif os.path.isdir(file_path): 
            shutil.rmtree(file_path)
    except Exception as e:
        print(e)