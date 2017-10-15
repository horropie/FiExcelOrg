import openpyxl
wb=openpyxl.Workbook()

wb.get_sheet_names()
sheet=wb.active
sheet.title
sheet.title="MyNewTitle"
wb.get_sheet_names()
wb.save("cr1.xlsx")
