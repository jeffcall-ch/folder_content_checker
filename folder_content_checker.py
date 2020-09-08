import pandas as pd
import pathlib
import logging
import os


df = pd.read_excel(r'\\AENBACIFS01.gaston.local\vol_tacdata1\data\Project Execution\0646 Marghera\07 GT Engineering\04 Piping and Arrangement (3D)\Pipe Support dwg - approved\0646 S85100AF002 01 - JO RedCorrex.xlsx')


df['Combined'] = df['SYSTEM'].astype(str)+'\\'+df['Part No.']
# fill nan with ''
df = df.fillna('')



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
   current_row_data = row_series[8]
   print (current_row_data)
   for current_file_path in listOfFiles:
       if current_file_path.find(current_row_data) > 0:
           print ('file found')
           # CHECK HOW TO WRITE RESULT TO DF!!
           df.at[index_label, 'File Downloaded'] = "YES"
           

print (df)


writer = pd.ExcelWriter(r'\\AENBACIFS01.gaston.local\vol_tacdata1\data\Project Execution\0646 Marghera\07 GT Engineering\04 Piping and Arrangement (3D)\Pipe Support dwg - approved\0646 S85100AF002 01 - JO RedCorrex.xlsx')
df.to_excel(writer, sheet_name='Sheet2')
writer.save