from odoo import fields, models


class L10nArArcaTaxConfig(models.Model):
    _name = 'l10n_ar.arca.tax.config'
    _description = 'Configuración de Impuestos de Percepción para Importación ARCA'
    _rec_name = 'company_id'

    # Modelo propio (no extiende res.company) para no tocar un modelo core
    # compartido por toda la instalación: se guarda un registro por compañía
    # con los impuestos de percepción del 'Libro de Compras' de ARCA fijados
    # a mano. Es la vía preferida (si están seteados, se usan tal cual, sin
    # buscar por nombre) para blindar el import automático contra
    # diferencias de formato en el nombre del impuesto (ej. espacios en
    # 'Perc IVA (3 %)') que rompían la búsqueda por nombre en producción.
    # Si no hay registro para una compañía, o el campo puntual está vacío,
    # el módulo sigue buscando por nombre/alícuota como hasta ahora.
    company_id = fields.Many2one(
        'res.company', string='Compañía', required=True, default=lambda self: self.env.company)

    tax_ganancias_id = fields.Many2one(
        'account.tax', string='Percepción de Ganancias',
        domain="[('company_id', '=', company_id), ('type_tax_use', '=', 'purchase')]",
        help="Impuesto a aplicar para la columna 'Importe de Per. o Pagos a Cta. de Otros Imp. "
             "Nac.' del Libro de Compras de ARCA. Si se deja vacío, se busca automáticamente por "
             "nombre ('Perc Gananc' / 'Percepción Ganancias').")
    tax_iibb_id = fields.Many2one(
        'account.tax', string='Percepción de IIBB',
        domain="[('company_id', '=', company_id), ('type_tax_use', '=', 'purchase')]",
        help="Impuesto a aplicar para la columna 'Importe de Percepciones de Ingresos Brutos' del "
             "Libro de Compras de ARCA. Si se deja vacío, se busca automáticamente por nombre "
             "('P. IIBB MZA' / 'Percepción IIBB Mendoza').")
    tax_perc_iva_21_id = fields.Many2one(
        'account.tax', string='Percepción de IVA - Neto 21%',
        domain="[('company_id', '=', company_id), ('type_tax_use', '=', 'purchase')]",
        help="Impuesto a aplicar para 'Percepciones/Pagos a Cta. de IVA' + 'Otros Tributos' del "
             "Libro de Compras cuando la línea de Neto Gravado de mayor monto del comprobante es al "
             "21% (o a una alícuota distinta de 10,5%). Si se deja vacío, se busca automáticamente "
             "por alícuota (impuesto de compras al 3%).")
    tax_perc_iva_105_id = fields.Many2one(
        'account.tax', string='Percepción de IVA - Neto 10,5%',
        domain="[('company_id', '=', company_id), ('type_tax_use', '=', 'purchase')]",
        help="Impuesto a aplicar para 'Percepciones/Pagos a Cta. de IVA' + 'Otros Tributos' del "
             "Libro de Compras cuando la línea de Neto Gravado de mayor monto del comprobante es al "
             "10,5%. Si se deja vacío, se busca automáticamente por alícuota (impuesto de compras "
             "al 1,5%).")

    _sql_constraints = [
        ('company_id_unique', 'unique(company_id)', 'Ya existe una configuración de impuestos ARCA para esta compañía.'),
    ]
