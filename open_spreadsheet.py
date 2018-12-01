import openpyxl

wb = openpyxl.load_workbook('grading.xlsx')
print(type(wb))