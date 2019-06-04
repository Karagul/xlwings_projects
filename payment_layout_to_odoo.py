import xlwings as xw
from xlwings import constants

if xw.Range('A1').value == 'Historial de Ordenes Enviadas':
    xw.Range('1:1').api.Delete(constants.DeleteShiftDirection.xlShiftUp)

original_ammount = xw.Range('H:H').value

original_sheet = xw.sheets[0]
original_sheet.name = 'ORIGINAL'

original_sheet.api.Copy(Before=original_sheet.api)
pagos_fusion_sheet = xw.sheets[0]
pagos_fusion_sheet.name = 'PAGOS FUSION'

pagos_fusion_sheet.api.Copy(Before=pagos_fusion_sheet.api)
pagos_fusion_polizas_sheet = xw.sheets[0]
pagos_fusion_polizas_sheet.name = 'PAGOS FUSION POLIZAS'

pagos_fusion_polizas_sheet.api.Copy(Before=pagos_fusion_polizas_sheet.api)
pagos_no_fusion = xw.sheets[0]
pagos_no_fusion.name = 'PAGOS NO DE FUSION'

pagos_no_fusion.api.Copy(Before=pagos_no_fusion.api)
devoluciones_sheet = xw.sheets[0]
devoluciones_sheet.name = 'DEVOLUCIONES'

cuadre_sheet = xw.sheets.add()
cuadre_sheet.name = 'CUADRE'
xw.Range('A1').value = 'MONTO REPORTE ORIGINAL'
xw.Range('B1').value = original_ammount


def create_payments_template():

    template_sheet = xw.sheets.add()
    template_sheet.name = 'TEMPLATE DE CARGA'

    xw.Range('A1').value = 'payment_type'
    xw.Range('B1').value = 'partner_type'
    xw.Range('C1').value = 'partner_id/id'
    xw.Range('D1').value = 'amount'
    xw.Range('E1').value = 'journal_id/id'
    xw.Range('F1').value = 'payment_date'
    xw.Range('G1').value = 'communication'
    xw.Range('H1').value = 'payment_method_id/id'
