import xlwings as xl
from xlwings import constants


def create_payments_template():
    template_sheet = xl.sheets.add()
    template_sheet.name = 'TEMPLATE DE CARGA'

    xl.Range('A1').value = 'payment_type'
    xl.Range('B1').value = 'partner_type'
    xl.Range('C1').value = 'partner_id/id'
    xl.Range('D1').value = 'amount'
    xl.Range('E1').value = 'journal_id/id'
    xl.Range('F1').value = 'payment_date'
    xl.Range('G1').value = 'communication'
    xl.Range('H1').value = 'payment_method_id/id'


def filter_column_hidden(col_num, filter_value):
    first_row = xl.Range('B2')
    print(first_row)
    last_row = xl.Range('B1').end('down')
    print(last_row)

    xl.Range("A1").api.AutoFilter(Field=col_num, Criteria1=filter_value, Operator=7)

    rng = xl.Range(first_row, last_row)

    hidden_rows = []

    for row in rng:
        if row.api.EntireRow.Hidden:
            hidden_rows.append(row)

    for row in reversed(hidden_rows):
        print(row)
        row.api.EntireRow.Delete()


def filter_column_active(col_num, filter_value):
    first_row = xl.Range('B2')
    print(first_row)
    last_row = xl.Range('B1').end('down')
    print(last_row)

    xl.Range("A1").api.AutoFilter(Field=col_num, Criteria1=filter_value, Operator=7)

    rng = xl.Range(first_row, last_row)

    active_rows = []

    for row in rng:
        if not row.api.EntireRow.Hidden:
            active_rows.append(row)

    for row in reversed(active_rows):
        print(row)
        row.api.EntireRow.Delete()

    xl.Range("A1").api.AutoFilter(Field=col_num)

def upload_payments_template():
    # Test
    return


# Removing first unnecessary column
if xl.Range('A1').value == 'Historial de Ordenes Enviadas':
    xl.Range('1:1').api.Delete(constants.DeleteShiftDirection.xlShiftUp)

# Getting the first Sheet
original_sheet = xl.sheets[0]
original_sheet.name = 'ORIGINAL'

# Auto fitting columns
original_sheet.autofit()

print('Filtering send orders')
filter_column_hidden(9, ['Liquidada', 'Traspaso Liquidado'])

first_row = xl.Range('H2')
last_row = xl.Range('H1').end('down')
rng = xl.Range(first_row, last_row)

original_amount = sum(rng.value)
print(original_amount)

# Copying the original sheet
original_sheet.api.Copy(Before=original_sheet.api)
pagos_fusion_sheet = xl.sheets[0]
pagos_fusion_sheet.name = 'PAGOS FUSION'

print('Filtering Company')
filter_column_hidden(5, ['MORALTA', 'SAN_FERNANDO', 'SIST_AGUA', 'TERRALTA', 'TRES VISTAS', 'GRUPO_FUSION'])

print('Filtering returns')
filter_column_active(7, ['Traspaso Final de Fondos'])

print('Filtering Salvatore Valassi')
filter_column_active(11, ['014180605382969691', '012180015223497953'])

first_row = xl.Range('H2')
last_row = xl.Range('H1').end('down')
rng = xl.Range(first_row, last_row)

fusion_amount = sum(rng.value)

original_sheet.api.Copy(Before=pagos_fusion_sheet.api)
pagos_fusion_polizas_sheet = xl.sheets[0]
pagos_fusion_polizas_sheet.name = 'PAGOS FUSION POLIZAS'

print('Filtering Salvatore Valassi')
filter_column_hidden(11, ['014180605382969691', '012180015223497953'])

first_row = xl.Range('H2')
last_row = xl.Range('H1').end('down')
rng = xl.Range(first_row, last_row)

try:
    manual_amount = sum(rng.value)

except:
    manual_amount = rng.value

original_sheet.api.Copy(Before=pagos_fusion_polizas_sheet.api)
pagos_no_fusion = xl.sheets[0]
pagos_no_fusion.name = 'PAGOS NO DE FUSION'

print('Filtering Company')
filter_column_active(5, ['MORALTA', 'SAN_FERNANDO', 'SIST_AGUA', 'TERRALTA', 'TRES VISTAS', 'GRUPO_FUSION'])

print('Filtering returns')
filter_column_active(7, ['Traspaso Final de Fondos'])

print('Filtering Salvatore Valassi')
filter_column_active(11, ['014180605382969691', '012180015223497953'])

first_row = xl.Range('H2')
last_row = xl.Range('H1').end('down')
rng = xl.Range(first_row, last_row)

not_fusion_amount = sum(rng.value)

original_sheet.api.Copy(Before=pagos_no_fusion.api)
devoluciones_sheet = xl.sheets[0]
devoluciones_sheet.name = 'DEVOLUCIONES'

print('Filtering returns')
filter_column_hidden(7, ['Traspaso Final de Fondos'])

first_row = xl.Range('H2')
last_row = xl.Range('H1').end('down')
rng = xl.Range(first_row, last_row)

returns_amount = sum(rng.value)

cuadre_sheet = xl.sheets.add()
cuadre_sheet.name = 'CUADRE'
xl.Range('A1').value = 'MONTO REPORTE ORIGINAL'
xl.Range('B1').value = original_amount

xl.Range('A3').value = 'PAGOS FUSION (SE CARGAN A ODOO)'
xl.Range('B3').value = fusion_amount
xl.Range('A4').value = 'PAGOS COMO PÓLIZA MANUAL'
xl.Range('B4').value = manual_amount
xl.Range('A5').value = 'PAGOS NO DE FUSIÓN'
xl.Range('B5').value = not_fusion_amount
xl.Range('A6').value = 'DEVOLUCIONES'
xl.Range('B6').value = returns_amount
xl.Range('A7').value = 'TOTAL'
xl.Range('B7').value = sum(xl.Range('B3:B6').value)

xl.Range('A9').value = 'DIFERENCIA'
xl.Range('B9').value = xl.Range('B1').value - xl.Range('B7').value
xl.sheets.active.autofit()


