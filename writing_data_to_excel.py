# this programme is for writing same data in every cell



import openpyxl


#(user can give any file path )
file_path = "python_data_driven_testing\data2.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook["Sheet1"]


# below logic write same data in every row and col
for r in range(1,6):
    for c in range(1,6):
        sheet.cell(r,c).value = "welcome"
  
 #important point :after writing data we must have to save workbook file otherwise empty file will show   
workbook.save(file_path)