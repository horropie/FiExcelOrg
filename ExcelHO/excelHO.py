import openpyxl
from operator import is_not
wb=openpyxl.load_workbook("Finanzen 2017.xlsx")
mysheet=wb.get_sheet_by_name("Juni")
#for col in mysheet.iter_cols(min_row=15, min_col=4, max_col=1, max_row=100):
#    for cell in col:
#        if mysheet[str(cell)].value==

def exists(it):
    return (it is not None)

'''
listb=[]
for j in range(15, 100):
    e=mysheet.cell(row=j, column=4)
    if e.value is not None:
        listb.append(e.value)

print(listb)
'''


s="=Summe("

for j in range(15, 100):
    e=mysheet.cell(row=j, column=4)
    mysheet["C8"]=s
    if e.value is "l":
        s+="C"+str(j)+","

s2=s
s2+=")"
mysheet["C8"]=s2

print(mysheet["C8"].value)
#now still need to add the celllocation and "Summe()"


wb.save('Finanzen 2017_Test.xlsx')
