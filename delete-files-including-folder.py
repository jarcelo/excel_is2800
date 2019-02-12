# This util will delete all contents in the directory including the folder itself

import shutil

directory = "C:\\Path\\folder"

print("Deleting all files in " + directory)

shutil.rmtree(directory)

print("Done")



