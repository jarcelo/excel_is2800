import os
import zipfile
from input_file import ExcelStagingDirectory

directory = ExcelStagingDirectory
zipExtension = ".zip"
xlsxExtension = ".xlsx"

os.chdir(directory)

# Extract the Excel file
for item in os.listdir(directory):
        if item.endswith(zipExtension):
                fileName = os.path.abspath(item)
                zipRef = zipfile.ZipFile(fileName)
                zipRef.extractall(directory)

# Open the excel files
for item in os.listdir(directory):
        if item.endswith(xlsxExtension):
                os.startfile(item)

                