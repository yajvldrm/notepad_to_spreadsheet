import openpyxl
wb =  openpyxl.load_workbook('grading.xlsx')
print(wb.get_sheet_names())
sheet = wb['Sheet1']
print(sheet)
print(type(sheet))
print(sheet.title)
anothersheet = wb.active
print(anothersheet)