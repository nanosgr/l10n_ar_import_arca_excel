import base64
import io
import openpyxl
from odoo.tests import common, tagged

@tagged('post_install', '-at_install')
class TestArcaImport(common.TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        # Create a journal
        cls.journal_purchase = cls.env['account.journal'].create({
            'name': 'Compras Test',
            'type': 'purchase',
            'code': 'CPT',
            'currency_id': cls.env.ref('base.ARS').id,
        })
        
        # Create a tax
        cls.tax_21 = cls.env['account.tax'].create({
            'name': 'IVA 21%',
            'type_tax_use': 'purchase',
            'amount': 21.0,
            'amount_type': 'percent',
            'country_id': cls.env.ref('base.ar').id,
        })

    def _create_dummy_excel(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        # Headers matching the logic
        headers = [
            'Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Número Hasta', 
            'Cód. Autorización', 'Tipo Doc. Emisor', 'Nro. Doc. Emisor', 'Denominación Emisor', 
            'Tipo Doc. Receptor', 'Nro. Doc. Receptor', 'Tipo Cambio', 'Moneda', 
            'Neto Grav. IVA 21%', 'IVA 21%', 'Imp. Total'
        ]
        ws.append(headers)
        # Row 1: Valid Import
        ws.append([
            '01/01/2026', '1 - Factura A', 1, 1234, 1234, 
            '12345678901234', 'CUIT', '30111111118', 'PROVEEDOR TEST', 
            'CUIT', '20111111112', 1, '$', 
            100.0, 21.0, 121.0
        ])
        
        output = io.BytesIO()
        wb.save(output)
        return base64.b64encode(output.getvalue())

    def test_import_process(self):
        excel_data = self._create_dummy_excel()
        
        wizard = self.env['l10n_ar.arca.import.wizard'].create({
            'file_data': excel_data,
            'filename': 'test.xlsx',
            'import_type': 'in_invoice',
            'default_journal_id': self.journal_purchase.id
        })
        
        # 1. Analyze
        wizard.action_analyze()
        
        self.assertEqual(len(wizard.line_ids), 1, "Should have 1 line")
        line = wizard.line_ids[0]
        self.assertEqual(line.status, 'ready', "Line should be ready")
        self.assertEqual(line.amount_total, 121.0)
        self.assertEqual(line.partner_name, 'PROVEEDOR TEST')
        
        # 2. Import
        wizard.action_import()
        
        # Verify invoice created
        move = line.move_id
        self.assertTrue(move, "Move should be created")
        self.assertEqual(move.amount_total, 121.0)
        self.assertEqual(move.partner_id.name, 'PROVEEDOR TEST')
        self.assertEqual(move.invoice_line_ids[0].tax_ids[0], self.tax_21)
