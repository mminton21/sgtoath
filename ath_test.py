import openpyxl

"""Program to take an excel file with rows that need to be concatenated selectively and execute on parameters"""

# Open up file in same folder and pick desired sheet.
ath = openpyxl.load_workbook('ath_data.xlsx')
sheet = ath["Worksheet"]

# loop through the cells from AH2 to AP7 to capture initial values in the field, then turn them into strings. Put
# each row in a list called cell_list, then append the entire list to cell_dict.
cell_list = []
cell_dict = []
for cell_values in sheet['AH2':'AP7']:
    for co in cell_values:
        if co.value:
            cell_list.append(str(co.value))
    cell_dict.append(cell_list)
    cell_list = []

# take the individual lists that were appended to cell_dict and append them to new_cell_dict as joined strings in the
# value^value^value format.
new_cell_dict = []
for contents in cell_dict:
    separator = "^"
    filler = separator.join(contents)
    new_cell_dict.append(filler)

# take the total number of rows that data exists in and place each set of lists created previously in the column
# specified, starting at row 2.
counter = len(new_cell_dict)
srow = 1
i = 0
while i < counter:
    str_of_list = str(new_cell_dict[i])
    sheet.cell(row=srow + 1, column=43).value = str_of_list
    i += 1
    srow += 1

ath.save('ath_data.xlsx')
