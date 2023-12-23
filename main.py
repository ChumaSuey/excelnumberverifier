import openpyxl
# This script just finds a number in a cell value in excel.

# Set the full path to the Excel file and read the excel file
filename = 'insert filename EXACT location'
wb = openpyxl.load_workbook(filename)

# The sheet will be the "1" or number one sheet name.

sheet = wb["1"]

# Search for the number
num_found = False
for row in sheet.iter_rows():
    for cell in row:
        if cell.value == 'XX-XX':
            num_found = True
            break


# Check if the number was found
if num_found:
    print('The number XX-XX is present in the Excel file {}.'.format(filename))
else:
    print('The number XX-XX is not present in the Excel file {}.'.format(filename))


#This script requires to tell the precise filename location (as in mainpractical.py)