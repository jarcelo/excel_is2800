# this will delete all contents in the directory including the folder itself

import shutil

directory = "Path"

print("Deleting all files in " + directory)

shutil.rmtree(directory)

print("Done")



