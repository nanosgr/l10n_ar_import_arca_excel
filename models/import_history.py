from odoo import models, fields, api

class L10nArArcaImportHistory(models.Model):
    _name = 'l10n_ar.arca.import.history'
    _description = 'Historial de Importación ARCA'
    _order = 'create_date desc'
    _inherit = ['mail.thread', 'mail.activity.mixin']

    name = fields.Char(string='Descripción', required=True)
    filename = fields.Char(string='Nombre Archivo')
    file_data = fields.Binary(string='Archivo Excel')
    user_id = fields.Many2one('res.users', string='Usuario', default=lambda self: self.env.user, readonly=True)
    import_type = fields.Selection([
        ('in_invoice', 'Facturas Recibidas'),
        ('out_invoice', 'Facturas Emitidas')
    ], string='Tipo', readonly=True)
    
    move_ids = fields.Many2many('account.move', string='Facturas Creadas', readonly=True)
    move_count = fields.Integer(string='Cant. Facturas', compute='_compute_move_count', store=True)
    
    @api.depends('move_ids')
    def _compute_move_count(self):
        for rec in self:
            rec.move_count = len(rec.move_ids)

    def action_view_moves(self):
        self.ensure_one()
        return {
            'name': 'Facturas Creadas',
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('id', 'in', self.move_ids.ids)],
            'context': {'create': False}
        }
