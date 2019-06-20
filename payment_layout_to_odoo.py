import xlwings as xl
from xlwings import constants
from Model import *


def create_payments_template():
    xl.sheets['PAGOS FUSION'].select()

    first_row = xl.Range('A2')
    last_row = xl.Range('A2').end('down')
    date = xl.Range(first_row, last_row)

    first_row = xl.Range('D2')
    last_row = xl.Range('D2').end('down')
    partner = xl.Range(first_row, last_row)

    first_row = xl.Range('C2')
    last_row = xl.Range('H1').end('down')
    print(last_row)
    rfc = xl.Range(first_row, last_row)
    rows = rfc.rows.count
    last_row = xl.Range('C'+str(rows+1))
    rfc = xl.Range(first_row, last_row)
    print(rows)
    print(rfc)

    first_row = xl.Range('G2')
    last_row = xl.Range('G2').end('down')
    description = xl.Range(first_row, last_row)

    first_row = xl.Range('H2')
    last_row = xl.Range('H2').end('down')
    amount = xl.Range(first_row, last_row)

    first_row = xl.Range('K2')
    last_row = xl.Range('K2').end('down')
    bank_account = xl.Range(first_row, last_row)

    try:
        xl.sheets['TEMPLATE DE CARGA'].select()
        xl.sheets['TEMPLATE DE CARGA'].clear()

    except:
        template_sheet = xl.sheets.add()
        template_sheet.name = 'TEMPLATE DE CARGA'

    xl.Range('A1').value = 'payment_type'
    xl.Range('B1').value = 'partner_type'
    xl.Range('C1').value = 'partner_id/database id'
    xl.Range('D1').value = 'amount'
    xl.Range('E1').value = 'journal_id/id'
    xl.Range('F1').value = 'payment_date'
    xl.Range('G1').value = 'communication'
    xl.Range('H1').value = 'payment_method_id/id'

    connection.startConnection()

    for i, row in enumerate(amount):
        # payment_type
        xl.Range((i + 2, 1)).value = 'Send Money'
        # partner_type
        xl.Range((i + 2, 2)).value = 'Proveedor'

        print('Searching by name: ' + partner[i].value)
        partner_id = connection.ODOO_OBJECT.execute_kw(
            connection.DATA,
            connection.UID,
            connection.PASS,
            'res.partner',
            'search',
            [[['name', '=', partner[i].value]]])

        if partner_id:
            # partner_id/id
            print('Partner ID: ' + str(partner_id[0]))
            xl.Range((i + 2, 3)).value = partner_id[0]

        elif bank_account[i].value:

            print('Searching by Bank Account: ' + bank_account[i].value)
            partner_id = connection.ODOO_OBJECT.execute_kw(
                connection.DATA,
                connection.UID,
                connection.PASS,
                'res.partner.bank',
                'search_read',
                [[['acc_number', '=', bank_account[i].value]]])

            if partner_id:
                # partner_id/id
                print('Partner ID: ' + str(partner_id[0]['partner_id'][0]))
                xl.Range((i + 2, 3)).value = partner_id[0]['partner_id'][0]

            elif rfc[i].value:

                print('Searching by RFC: ' + rfc[i].value)

                partner_id = connection.ODOO_OBJECT.execute_kw(
                    connection.DATA,
                    connection.UID,
                    connection.PASS,
                    'res.partner',
                    'search',
                    [[['vat', '=', rfc[i].value]]])

                if partner_id:
                    # partner_id/id
                    print('Partner ID: ' + str(partner_id[0]))
                    xl.Range((i + 2, 3)).value = partner_id[0]

                else:
                    xl.Range((i + 2, 3)).value = 'NO EXISTE EL PROVEEDOR'

            else:
                xl.Range((i + 2, 3)).value = 'NO EXISTE EL PROVEEDOR'

        # amount
        xl.Range((i+2, 4)).value = row.value
        # journal_id/id
        xl.Range((i + 2, 5)).value = '__export__.account_journal_21_5b9100d0'
        # payment_date
        xl.Range((i + 2, 6)).value = datetime.datetime.strptime(date[i].value, '%d/%m/%Y')
        # communication
        xl.Range((i + 2, 7)).value = description[i].value
        # payment_method_id/id
        xl.Range((i + 2, 8)).value = 'account.account_payment_method_manual_out'


    xl.Range('A:H').autofit()
    xl.sheets['TEMPLATE DE CARGA'].api.Move(Before=xl.sheets[0].api)

# Metodo que elimina las columnas escondidas
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


# Metodo que elimina las columnas visibles
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


def generate_payments_base():
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

    print(first_row)
    print(last_row)
    print(rng.value)

    manual_amount = 0

    if rng.rows.count < 60000:
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

    try:
        returns_amount = sum(rng.value)
    except:
        returns_amount = rng.value

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


def upload_payments_template():
    # Test
    return

connection = Connection()

#create_payments_template()

generate_payments_base()
