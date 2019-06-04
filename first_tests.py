import xlwings as xw

#Creation of new workbook
#wb = xw.Book()

xw.Range('A1').value = 'something'

ws1 = xw.sheets[0]

print(ws1.name)