import pandas as pd
import pathlib
import logging
import os


df = pd.read_excel(r'\\AENBACIFS01.gaston.local\vol_tacdata1\data\Project Execution\0646 Marghera\07 GT Engineering\04 Piping and Arrangement (3D)\Pipe Support dwg - approved\0646 S85100AF002 01 - JO RedCorrex.xlsx')


df['Combined'] = df['SYSTEM'].astype(str)+'\\'+df['Part No.']

print (df)

path_to_dwgs = r'\\AENBACIFS01.gaston.local\vol_tacdata1\data\Project Execution\0646 Marghera\07 GT Engineering\04 Piping and Arrangement (3D)\Pipe Support dwg - approved'

def getListOfFiles(dirName):
    # create a list of file and sub directories 
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    allFiles = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)
                
    return allFiles     

listOfFiles = getListOfFiles(path_to_dwgs)

# for elem in listOfFiles:
#        print(elem)

for (index_label, row_series) in df.iterrows():
   print('Row Index label : ', index_label)
   # print('Row Content as Series : ', row_series.values)
   print (row_series[7])