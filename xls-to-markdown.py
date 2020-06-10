import pandas as pd
import os
import glob


def maxlen(sheet):                                   # this function creates a list of the max lengths of data etered in each row
    dim = sheet.shape
    maxsize = [' '] * dim[1]
    for i in range(dim[1]):
        maxsize[i] = len(max(sheet.iloc[:,i], key = len))
    return maxsize


def xls2md(sheet, file_name, sheet_name):

    dim1 = sheet.shape                           # to find out dimensions of the sheet

    for i in range(dim1[0]):
        for j in range(dim1[1]):
            sheet.iloc[i, j] = str(sheet.iloc[i, j])     # converts all data in the spreadsheet to string

    indices = sheet.columns                    #gets a list of names of column headers in a sheet
    maxsize = maxlen(sheet)
    dest_path = os.getcwd() + '\\'
    dest_path = dest_path + 'markdown\\'
    dest_path = dest_path + file_name + '_' + sheet_name + '.md'    #destination file for markdown
    f = open(dest_path, 'w')

    for i in range(dim1[1]):                                  #this loop determines number of spaces to be put for the column headers
        if maxsize[i] < len(indices[i]):
            f.write('| ' + indices[i] + ' ')
        else:
            f.write('| ' + indices[i] + ' '*(maxsize[i] - len(indices[i]) + 1 ) )
    f.write('|')
    f.write("\n")

    for i in range(dim1[1]):             # this loop writes separator for the column headers and the data 
        if len(indices[i]) >= maxsize[i]:
            f.write('|' + '-'*(len(indices[i]) + 2))
        else:
            f.write('|' + '-'*(maxsize[i] + 2))
    f.write('|')
    f.write("\n")

    for i in range(dim1[0]):           #this loop writes the data into the file withe adequate spaces
        for j in range(dim1[1]):
            f.write('| ' + sheet.iloc[i,j])
            if maxsize[j] < len(indices[j]):
                f.write(' ' * (len(indices[j]) - len(sheet.iloc[i, j]) + 1))
            else:
                f.write(' ' * (maxsize[j] - len(sheet.iloc[i, j]) + 1))
        f.write('|')
        f.write("\n")


cur_path = os.getcwd()                                            # gets the path where the python file is located
path = cur_path + r'\excel'
files = [f for f in glob.glob(path + '**/*.xlsx', recursive=True)] # gets the paths for all the excel files stored in 'excel' folder 
filename = []
lenpath = len(path) + 1

for f in files:
    filename.append(f[lenpath:len(f) - 5])          #gets the filenames of all the files in the 'excel' folder

for i in range(len(filename)):
    spreadsheet = pd.ExcelFile(files[i])            #opens all the excel files one by one
    spreadsheet_sheets = spreadsheet.sheet_names    #gets the list of names of all the sheets in an excel file
    for j in range(len(spreadsheet_sheets)):            
        sheet = pd.read_excel(spreadsheet, spreadsheet_sheets[j]) #opens all the sheets in an excel file sequentially
        sheetname = spreadsheet_sheets[j]          
        xls2md(sheet, filename[i], sheetname)           #the function that converts the excel spreadsheet to md

print("All Files have been converted to Markdown")






