# This util will delete all contents in the directory including the folder itself

import shutil

directory = "C:\\Users\\Master Joe\\Desktop\\case3"

print("Deleting all files in " + directory)

shutil.rmtree(directory)

print("Done")



