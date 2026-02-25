# %%
import sys
import os
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string
import pandas as pd


import matplotlib.pyplot as mpl
from pathlib import Path

#arg format: inputdirectory cellnum imagefilename
cellNum = sys.argv[1]   #cell number (example: "A2", is this a string?)
imageName = sys.argv[2] #image name
fileList = sys.argv[3:] #slice of srd index over (list of file directories)
#xlfiles = Path(folderPath).iterdir()

#if last char of folderPath is /, change nothing, otherwise append it to end

#Convert the column letter to a number. 
#A -> 1
#AA -> 27
def column_letter_to_number(col):
    col = col.upper()
    result = 0
    for char in col:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


try:
    coordinate_from_string(cellNum)
except Exception:
    print("Cell is invalid! Try again. ")
    sys.exit(0)
index = 0
emptyCells = 0
stringCells = 0
errorCells = 0
figure, axes = mpl.subplots()
values = []
x_names = []

    #for xlfile in sorted(list(xlfiles)):

def grabCell(xlfile):
    global index
    global emptyCells
    global stringCells
    global errorCells
    #currentFile = load_workbook(filename = (xlfile), data_only = True)
    #currentFile = currentFile.active 
    
    
    currentFile = pd.read_excel(xlfile, sheet_name=0, header=None) #Creates a pandas DataFrame
    #cellNum[0] 'A', the column.
    #cellNum[1] '1', the row
    #iloc[row,col]
    col_letters, row = coordinate_from_string(cellNum)
    row_index = row - 1
    col_index = column_letter_to_number(col_letters) - 1
    print(f"Row: {row_index} Col: {col_index}")

    cellValue = currentFile.iat[row_index, col_index]
    print(cellValue)
    if(cellValue == None):
        print(f"Cell {cellNum} in file {xlfile} is empty")
        emptyCells += 1
        return

    if(isinstance(cellValue,str)):
        print(f"File {xlfile} contains string data OR an error code!")
        #print(currentFile[cellNum].data_type)
        stringCells += 1
        return
    
    if(isinstance(cellValue,int) or isinstance(cellValue,float)):
        values.append(cellValue) #access the single chosen cell per file
        x_names.append(os.path.basename(xlfile))
    else:
        return
    
    
    print("plotting sheet " + str(xlfile)) 


for xlfile in fileList:
    if(os.path.exists(xlfile) == False):  
        sys.exit(f"{xlfile} can't be found!")


def fileLoop(fileList):
    for xlfile in fileList:
        if(os.path.isfile(xlfile)):
        # print(f"{xlfile} is a file")
            #If the fsile doesn't end in "xlsx", skip this iteration.
            if (str(xlfile)[-4:] != "xlsx"):
                print("non xlsx file detected!")
                continue
            else:
                grabCell(xlfile)

        elif(os.path.isdir(xlfile)):
            #print(f"{xlfile} is a directory!")
            xlList = Path(xlfile).iterdir()
            for elem in sorted(list(xlList)):
                #If there is a directory inside a directory, list those too
                if(os.path.isdir(elem)):
                    elemFolder = Path(elem).iterdir()
                    fileLoop(sorted(list(elemFolder)))
                else:
                    if (str(elem)[-4:] != "xlsx"):
                        print(f"non xlsx file detected! ({elem})")
                        continue
                    else:
                        grabCell(elem)
        else:
            print("Neither")

fileLoop(fileList)
       
print(x_names)
print(f"Values: {values}")
axes.plot(x_names,values)  
print("generating plot")
mpl.savefig(imageName + ".png")
print("plot generated")
print(f"Image saved as {os.path.abspath(imageName)}.png")
if(stringCells > 0):
    print(f"{stringCells} cells have string data and were excluded from graph")
if(emptyCells > 0):
    print(f"There were {emptyCells} empty cells detected.")
if (errorCells > 0):
    print(f"{errorCells} cells contained errors")

# %%
