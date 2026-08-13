from odoo import models, fields, api

class L10nArImportArcaPartnerFilter(models.TransientModel):
    _name = 'l10n_ar.arca.import.partner.filter'
    _description = 'Filtro de Partner para Importación ARCA'
    _order = 'amount_total desc'

    wizard_id = fields.Many2one('l10n_ar.arca.import.wizard', string='Wizard', ondelete='cascade')
    selected = fields.Boolean(string='Seleccionar', default=False)
    

    
    name = fields.Char(string='Razón Social (Excel)', readonly=True)
    cuit = fields.Char(string='CUIT', readonly=True)
    partner_id = fields.Many2one('res.partner', string='Contacto Relacionado', readonly=True)
    
    invoice_count = fields.Integer(string='# Comprobantes', readonly=True)
    amount_total = fields.Monetary(string='Total', currency_field='currency_id', readonly=True)
    currency_id = fields.Many2one('res.currency', string='Moneda')
    
    @api.onchange('selected')
    def _onchange_selected(self):
        # Trigger wizard update when selected changes
        # This might not be enough for full refresh, we likely need an action or depends
        pass
