import pandas as pd
import os
import glob


def maxlen(sheet):
    dim = sheet.shape
    maxsize = [' '] * dim[1]
    for i in range(dim[1]):
        maxsize[i] = len(max(sheet.iloc[:,i], key = len))
    return maxsize


def xls2md(sheet, file_name, sheet_name):

    dim1 = sheet.shape

    for i in range(dim1[0]):
        for j in range(dim1[1]):
            sheet.iloc[i, j] = str(sheet.iloc[i, j])

    indices = sheet.columns
    maxsize = maxlen(sheet)
    dest_path = os.getcwd() + '\\'
    dest_path = dest_path + 'markdown\\'
    dest_path = dest_path + file_name + '_' + sheet_name + '.md'
    f = open(dest_path, 'w')

    for i in range(dim1[1]):
        if maxsize[i] < len(indices[i]):
            f.write('| ' + indices[i] + ' ')
        else:
            f.write('| ' + indices[i] + ' '*(maxsize[i] - len(indices[i]) + 1 ) )
    f.write('|')
    f.write("\n")

    for i in range(dim1[1]):
        if len(indices[i]) >= maxsize[i]:
            f.write('|' + '-'*(len(indices[i]) + 2))
        else:
            f.write('|' + '-'*(maxsize[i] + 2))
    f.write('|')
    f.write("\n")

    for i in range(dim1[0]):
        for j in range(dim1[1]):
            f.write('| ' + sheet.iloc[i,j])
            if maxsize[j] < len(indices[j]):
                f.write(' ' * (len(indices[j]) - len(sheet.iloc[i, j]) + 1))
            else:
                f.write(' ' * (maxsize[j] - len(sheet.iloc[i, j]) + 1))
        f.write('|')
        f.write("\n")


cur_path = os.getcwd()
path = cur_path + r'\excel'
files = [f for f in glob.glob(path + '**/*.xlsx', recursive=True)]
filename = []
lenpath = len(path) + 1

for f in files:
    filename.append(f[lenpath:len(f) - 5])

for i in range(len(filename)):
    spreadsheet = pd.ExcelFile(files[i])
    spreadsheet_sheets = spreadsheet.sheet_names
    for j in range(len(spreadsheet_sheets)):
        sheet = pd.read_excel(spreadsheet, spreadsheet_sheets[j])
        sheetname = spreadsheet_sheets[j]
        xls2md(sheet, filename[i], sheetname)

print("All Files have been converted to Markdown")






