import xlwings as xw

#Creation of new workbook
#wb = xl.Book()

print(xw.Range('B1').number_format)

ws1 = xw.sheets[0]

