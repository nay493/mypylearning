#!/usr/bin/py3

import xlrd
fil = "/home/bonchuda/Desktop/formating sheet.xlsx"
workbook = xlrd.open_workbook(fil)
print("Workbook: ",workbook)
print(dir(workbook))
print("sheet name= ",dir(workbook.sheet_names))

sheet1 = workbook.sheet_by_index(1)
print("(0,0)=",sheet1.cell(0,0).value)
i = 0
j = 0
while i <=10207:
	print(sheet1.cell(i,4).value)
	i = i + 1
	while j <= 10207:
		if sheet1.cell(i,5).value == sheet1.cell(j,5).value:
			print("i,j= ",i,j)
		j = j + 1
exit()
