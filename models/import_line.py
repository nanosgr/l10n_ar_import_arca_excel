from odoo import models, fields, api
import json

class L10nArArcaImportLine(models.TransientModel):
    _name = 'l10n_ar.arca.import.line'
    _description = 'Línea de Importación ARCA'

    wizard_id = fields.Many2one('l10n_ar.arca.import.wizard', string='Asistente')
    
    # Datos Crudos del Excel
    date = fields.Date(string='Fecha')
    document_type_name = fields.Char(string='Tipo Doc.')
    point_of_sale = fields.Integer(string='Pto. Venta')
    number = fields.Integer(string='Número')
    cuit = fields.Char(string='CUIT')
    afip_auth_code = fields.Char(string='Código Autorización')
    partner_name = fields.Char(string='Razón Social')
    amount_total = fields.Monetary(string='Total', currency_field='currency_id')
    currency_id = fields.Many2one('res.currency', string='Moneda')
    
    # Datos Resueltos
    partner_id = fields.Many2one('res.partner', string='Contacto Encontrado')
    journal_id = fields.Many2one('account.journal', string='Diario Destino')
    move_id = fields.Many2one('account.move', string='Factura Existente')
    
    # Estado y Errores
    status = fields.Selection([
        ('ready', 'Listo'),
        ('exists', 'Ya existe'),
        ('error', 'Error / Bloqueante')
    ], string='Estado', default='ready')
    
    error_desc = fields.Html(string='Detalle Error')
    
    # Data Dump para creación posterior
    invoice_values = fields.Text(string='Valores JSON')
    
    to_import = fields.Boolean(string='Importar', default=False)

    preview_desc = fields.Char(string='Vista Previa')

    source_format = fields.Selection([
        ('mis_comprobantes', 'Mis Comprobantes (Excel)'),
        ('libro_compras', 'Libro de Compras (CSV)'),
    ], string='Origen del Dato', default='mis_comprobantes')

    # Percepciones e impuestos discriminados (solo presentes en 'Libro de Compras')
    credito_fiscal_computable = fields.Float(
        string='Crédito Fiscal Computable',
        help="Informativo: porción del IVA total que ARCA considera computable como crédito fiscal. No se suma al total del comprobante.")
    perc_otros_imp_nacionales = fields.Float(string='Perc./Pagos Cta. Otros Imp. Nac.')
    perc_iibb = fields.Float(string='Percepciones IIBB')
    imp_municipales = fields.Float(string='Impuestos Municipales')
    perc_iva = fields.Float(string='Perc./Pagos Cta. IVA')
    imp_internos = fields.Float(string='Impuestos Internos')

    @api.depends('status')
    def _compute_display_name(self):
        for rec in self:
            rec.display_name = f"{rec.document_type_name} {rec.point_of_sale}-{rec.number}"

    def action_open_error(self):
        # Placeholder por si queremos abrir detalle
        pass

    def action_toggle_import(self):
        for line in self:
            if line.status == 'ready':
                line.to_import = not line.to_import
        
        # Reload the wizard view to reflect changes
        # Use wizard_id to target the correct record
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'l10n_ar.arca.import.wizard',
            'res_id': self[0].wizard_id.id,
            'view_mode': 'form',
            'target': 'new',
        }
