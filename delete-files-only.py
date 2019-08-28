# This util will delete the files recursively within the folder

import os, shutil
from input_file import TargetTempDirectory

folder = TargetTempDirectory

for the_file in os.listdir(folder):
    file_path = os.path.join(folder, the_file)
    try:
        if os.path.isfile(file_path):
            os.unlink(file_path)
        elif os.path.isdir(file_path): 
            shutil.rmtree(file_path)
    except Exception as e:
        print(e)