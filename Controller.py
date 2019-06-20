from Model import *
from tkinter import *
import random
from tkinter.filedialog import askopenfilename
from xml.dom import minidom
import re


def main():
    conn = Connection()
    conn.startConnection()

    # testPartnerCreation(connection=conn)
    # testPurchaseOrderCreation(connection=conn)
    # testInvoiceCreation(connection=conn)
    get_invoices_to_approve(connection=conn)


def get_invoices_to_approve(connection):

    department_id = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'hr.department',
        'search',
        [[['name', '=', 'Construcción de Obra']]])

    invoice_ids = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'account.invoice',
        'search_read',
        [[['date_invoice', '=', '2019-06-06'], ['department_id', '=', department_id]]])
    # datetime.date.today()

    print(invoice_ids)

    for invoice in invoice_ids:
        print(invoice['partner_id'])
        print(invoice['reference'])
        print(invoice['amount_total'])
        print(invoice['residual'])
    '''
    [record] = connection.ODOO_OBJECT.execute_kw(
        connection.DATA,
        connection.UID,
        connection.PASS,
        'account.invoice',
        'read',
        [invoice_ids],
        {'fields': ['partner_id', 'reference', 'amount_total', 'residual']})

    if record:

        print(record)

    else:
        messagebox.showinfo(title="Partner Check",
                            message="There are no invoices to approve for today")
    '''

if __name__ == '__main__':

        main()
