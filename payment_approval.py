import xlwings as xl
from Model import *
from datetime import date
from tkinter import messagebox


def get_payments_to_approve():
    connection = Connection()
    connection.startConnection()

    department_id = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'hr.department',
        'search',
        [[['name', '=', 'Construcción de Obra']]])

    invoices = []

    invoices = invoices + connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'account.invoice',
        'search_read',
        [[['date_invoice', '=', str(date.today())], ['department_id', '=', department_id], ['state', '!=', 'approved_by_manager']]])

    department_id = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'hr.department',
        'search',
        [[['name', '=', 'Compras']]])

    invoices = invoices + connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'account.invoice',
        'search_read',
        [[['date_invoice', '=', str(date.today())], ['department_id', '=', department_id], ['state', '!=', 'approved_by_manager']]])

    department_id = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'hr.department',
        'search',
        [[['name', '=', 'Proyectos']]])

    invoices = invoices + connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'account.invoice',
        'search_read',
        [[['date_invoice', '=', str(date.today())], ['department_id', '=', department_id], ['state', '!=', 'approved_by_manager']]])

    return invoices


def fill_excel_sheet():

    invoices = get_payments_to_approve()
    if invoices:
        # COLUMN 1
        xl.Range('A1').value = 'ID DOC ODOO'
        # COLUMN 2
        xl.Range('B1').value = 'DESARROLLO'
        # COLUMN 3
        xl.Range('C1').value = 'SOLICITADO POR'
        # COLUMN 3
        xl.Range('D1').value = 'PROVEEDOR'
        # COLUMN 4
        xl.Range('E1').value = 'PRESUPUESTO'
        # COLUMN 5
        xl.Range('F1').value = 'DESCRIPCIÓN'
        # COLUMN 6
        xl.Range('G1').value = 'ORDEN DE COMPRA'
        # COLUMN 7
        xl.Range('H1').value = 'FACTURA'
        # COLUMN 8
        xl.Range('I1').value = 'MONTO'
        # COLUMN 9
        xl.Range('J1').value = 'FECHA DE FACTURA'
        # COLUMN 10
        xl.Range('K1').value = 'DÍAS TRANSCURRIDOS'
        # COLUMN 11
        xl.Range('L1').value = 'ESTADO'
        # COLUMN 12
        xl.Range('M1').value = 'MONTO AUTORIZADO'

        connection = Connection()
        connection.startConnection()

        for i, invoice in enumerate(invoices):
            print(invoice)
            if invoice['state'] == 'payment_request' or invoice['state'] == 'approved_by_leader' or invoice['state'] == 'open':
                xl.Range((i + 2, 1)).value = invoice['id']
                if invoice['account_analytic_id']:
                    xl.Range((i + 2, 2)).value = invoice['account_analytic_id'][1]
                xl.Range((i + 2, 3)).value = invoice['create_uid'][1]
                xl.Range((i + 2, 4)).value = invoice['partner_id'][1]

                invoice_line = connection.ODOO_OBJECT.execute_kw(
                    connection.DATA,
                    connection.UID,
                    connection.PASS,
                    'account.invoice.line',
                    'search_read',
                    [[['id', '=', invoice['invoice_line_ids']]]])

                xl.Range((i + 2, 5)).value = invoice_line[0]['product_id'][1]
                xl.Range((i + 2, 6)).value = invoice_line[0]['name'][0:40]
                if invoice['origin']:
                    xl.Range((i + 2, 7)).value = invoice['origin']
                if invoice['reference']:
                    xl.Range((i + 2, 8)).value = invoice['reference']
                if invoice['residual']:
                    xl.Range((i + 2, 9)).value = invoice['residual']
                else:
                    xl.Range((i + 2, 9)).value = invoice['amount_total']
                if invoice['x_invoice_date_sat']:
                    xl.Range((i + 2, 10)).value = invoice['x_invoice_date_sat']
                    xl.Range((i + 2, 11)).value = (datetime.datetime.now() - datetime.datetime.strptime(str(invoice['x_invoice_date_sat']), '%Y-%m-%d')).days
                xl.Range((i + 2, 12)).value = invoice['state']
                xl.Range((i + 2, 13)).value = xl.Range((i + 2, 9)).value
            else:
                continue

        xl.Range('A:M').autofit()
    else:
        messagebox.showinfo(message="No hay Facturas por Autorizar", title="Información Importante")


def upload_authorized_payments():
    first_row = xl.Range('A2')
    last_row = xl.Range('A1').end('down')
    rng = xl.Range(first_row, last_row)

    connection = Connection()
    connection.startConnection()

    for i, row in enumerate(rng):
        try:
            if xl.Range((i + 2, 9)).value and xl.Range((i + 2, 13)).value <= xl.Range((i + 2, 9)).value:
                if xl.Range((i + 2, 12)).value == "payment_request" or xl.Range((i + 2, 12)).value == "approved_by_leader":
                    connection.ODOO_OBJECT.execute_kw(
                        connection.DATA,
                        connection.UID,
                        connection.PASS,
                        'account.invoice',
                        'write',
                        [[int(row.value)], {'amount_authorized': xl.Range((i + 2, 13)).value, 'state': 'approved_by_manager'}])
                    print('successful upload of: ' + str(row.value) + ' Authorized Amount: $' + str(
                        xl.Range((i + 2, 13)).value))
                elif xl.Range((i + 2, 12)).value == "open":
                    connection.ODOO_OBJECT.execute_kw(
                        connection.DATA,
                        connection.UID,
                        connection.PASS,
                        'account.invoice',
                        'write',
                        [[int(row.value)],
                         {'amount_authorized': xl.Range((i + 2, 13)).value}])
                    print('successful upload of: ' + str(row.value) + ' Authorized Amount: $' + str(
                        xl.Range((i + 2, 13)).value))
            else:
                connection.ODOO_OBJECT.execute_kw(
                    connection.DATA,
                    connection.UID,
                    connection.PASS,
                    'account.invoice',
                    'write',
                    [[int(row.value)], {'state': 'payment_rejected'}])
                print('Payment Rejected of ' + str(xl.Range((i + 2, 9)).value))

        except:

            print('Error in line: ' + str(i))

#fill_excel_sheet()
#upload_authorized_payments()
