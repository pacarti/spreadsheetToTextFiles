def column_max_row(column_name):
    return max((c.row for c in sheet[column_name] if c.value is not None))

import os, openpyxl

os.chdir(os.path.dirname(os.path.abspath(__file__)))

wb = openpyxl.load_workbook('resultSpreadsheetFromTxtFiles.xlsx', data_only=True)

sheet = wb.active


# Check through a column for the max filled row

# TODO: Step2: Create a loop from below to make one file per column 

textFileA = open('textFileA.txt', 'w')

for i in range(1, column_max_row('A') + 1):
    # textFileA = open('textFileA.txt', 'a')
    # If there are an empty line between values, add it into the text file and continue through loop:
    '''
    if sheet['A' + str(i)].value == None:
        textFileA.write('')
        continue
    '''
    cellValue = sheet['A' + str(i)].value
    print(cellValue)

    if cellValue is not None:
        textFileA.write(cellValue + '\n')
    else:
        textFileA.write('\n')    
    
    # textFileA.write('\n')
    # textFileA.close()

textFileA.close()

# print(column_max_row('C'))

