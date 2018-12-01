import openpyxl

wb = openpyxl.load_workbook('grading.xlsx')
sheet = wb['Sheet1']
sheet['D2'].value = 40
wb.save('grading.xlsx')