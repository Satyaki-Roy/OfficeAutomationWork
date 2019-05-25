import xlrd
import xlwt
import os

# Deletes the Output.xls file if it already exists
os.remove("Output.xls")

# Handle for InventoryListFromSCCM.xlsx file
excel_sheet1 = "InventoryListFromSCCM.xlsx"

# Handle for SolarWindsList.xlsx file
excel_sheet2 = "SolarWindsList.xlsx"

# Handle to open the InventoryListFromSCCM.xlsx file in read mode
book1 = xlrd.open_workbook(excel_sheet1)

# Handle to open the SolarWindsList.xlsx file in read mode
book2 = xlrd.open_workbook(excel_sheet2)

# Handle to grab the first sheet of the file InventoryListFromSCCM.xlsx by index 0 (in python indexing starts from zero)
first_sheet = book1.sheet_by_index(0)

# Handle to grab the first sheet of the file SolarWindsList.xlsx by index 0 (in python indexing starts from zero)
second_sheet = book2.sheet_by_index(0)

# Handle to grab the column zero of SolarWindsList.xlsx
x = second_sheet.col_values(0)

# Handle to grab the column zero of InventoryListFromSCCM.xlsx
z = first_sheet.col_values(0)

# Declaring "y" as a list
# Initializing count
y = []
count = 0

# Converting the column 1 of SolarWindsList.xlsx (Eg:- mr01mj16dde.accudyneindustries.com to MR01MJ16DDE)
# and representing it as a list
for i in x:
    y.append(str(x[count]).replace('.accudyneindustries.com', '').upper())
    count += 1

# Handle for the Output.xls file in write mode
# Handle for the worksheet and worksheet name is 'My Worksheet'
workbook = xlwt.Workbook(encoding='ascii')
worksheet = workbook.add_sheet('My Worksheet')

# Variable to point the row in the Output.xls file
row_number_output = 0

for i in y:
    # Variable to point the row in the InventoryListFromSCCM.xlsx file
    row_number_inventory_list = 0

    # Logic to check the column ZERO of, InventoryListFromSCCM.xlsx and SolarWindsList.xlsx
    # also to write in Output.xls file
    for j in z:
        if i == j:
            temp0 = str(first_sheet.row_values(row_number_inventory_list)[0])
            temp1 = str(first_sheet.row_values(row_number_inventory_list)[1])
            temp2 = str(first_sheet.row_values(row_number_inventory_list)[2])
            worksheet.write(row_number_output, 0, label=temp0)
            worksheet.write(row_number_output, 1, label=temp1)
            worksheet.write(row_number_output, 2, label=temp2)
            row_number_output += 1
        else:
            pass
        row_number_inventory_list += 1

# Saving the workbook as Output.xls
workbook.save('Output.xls')
