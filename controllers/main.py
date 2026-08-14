import base64
import hmac
import logging

import odoo
from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)


class ArcaImportController(http.Controller):
    """ Endpoint HTTP para que un proceso externo (ej. el scraper diario de
    ARCA) suba el CSV del 'Libro de Compras' y dispare la importación
    automática, sin intervención de un usuario en la interfaz. """

    @http.route('/l10n_ar_arca_import/upload', type='http', auth='none', csrf=False, methods=['POST'])
    def upload_libro_compras(self, **kw):
        expected_token = request.env['ir.config_parameter'].sudo().get_param(
            'l10n_ar_import_arca_excel.api_token', '')
        token = kw.get('token', '')
        if not expected_token or not hmac.compare_digest(token, expected_token):
            return request.make_json_response({'success': False, 'error': 'Token inválido.'}, status=403)

        upload = request.httprequest.files.get('file')
        if not upload:
            return request.make_json_response(
                {'success': False, 'error': "Falta el archivo 'file'."}, status=400)

        try:
            company_id = int(kw.get('company_id', ''))
        except (TypeError, ValueError):
            return request.make_json_response(
                {'success': False, 'error': "Falta o es inválido el parámetro 'company_id'."}, status=400)

        # 'auth=none' deja request.env sin un usuario real vinculado
        # (env.uid vacío), lo que rompe cualquier cómputo interno de Odoo que
        # haga self.env.user (ej. account_accountant._compute_signing_user).
        # Se rebindea el entorno a un usuario real (Superusuario) para que
        # todo el flujo se comporte como si lo ejecutara un usuario logueado.
        env = request.env(user=odoo.SUPERUSER_ID)

        company = env['res.company'].browse(company_id)
        if not company.exists():
            return request.make_json_response(
                {'success': False, 'error': f"La compañía {company_id} no existe."}, status=400)

        filename = upload.filename or 'libro_compras.csv'
        file_content = upload.read()

        wizard = env['l10n_ar.arca.import.wizard'].with_company(company).create({
            'file_data': base64.b64encode(file_content),
            'filename': filename,
            'company_id': company.id,
            'import_type': 'in_invoice',
        })

        try:
            wizard.action_analyze()
        except Exception as e:
            _logger.exception("Error al analizar '%s' (Libro de Compras, API)", filename)
            self._send_notification_email(env, company, filename, error=str(e))
            return request.make_json_response({'success': False, 'error': f"Error al analizar: {e}"})

        # Se guarda el estado post-análisis para poder distinguir, después de
        # importar, entre 'recién creado' y 'ya existía' (ambos terminan en
        # status='exists' una vez corrido action_import).
        pre_status = {line.id: line.status for line in wizard.line_ids}

        try:
            wizard.action_import()
        except Exception as e:
            _logger.exception("Error al importar '%s' (Libro de Compras, API)", filename)
            self._send_notification_email(env, company, filename, error=str(e))
            return request.make_json_response({'success': False, 'error': f"Error al importar: {e}"})

        summary = self._build_summary(wizard, pre_status)
        self._send_notification_email(env, company, filename, summary=summary)
        _logger.info(
            "Importación ARCA (API) '%s' compañía %s: %s nuevas, %s ya existentes, %s con error",
            filename, company.name, summary['imported_count'], summary['duplicates_count'],
            summary['failed_count'])
        return request.make_json_response({'success': True, 'summary': summary})

    def _build_summary(self, wizard, pre_status):
        imported = []
        failed = []
        duplicates_count = 0

        for line in wizard.line_ids:
            pre = pre_status.get(line.id)
            if pre == 'ready':
                if line.status == 'exists' and line.move_id:
                    imported.append(line)
                else:
                    failed.append(line)
            elif pre == 'exists':
                duplicates_count += 1
            elif pre == 'error':
                failed.append(line)

        def _line_info(line):
            return {
                'partner_name': line.partner_name,
                'document': line.preview_desc,
                'amount_total': line.amount_total,
                'error': (line.error_desc or '').replace('<br/>', ' | '),
            }

        return {
            'total_lines': len(wizard.line_ids),
            'imported_count': len(imported),
            'duplicates_count': duplicates_count,
            'failed_count': len(failed),
            'errors': [_line_info(line) for line in failed],
        }

    def _send_notification_email(self, env, company, filename, summary=None, error=None):
        to_email = env['ir.config_parameter'].get_param('l10n_ar_import_arca_excel.notification_email')
        if not to_email:
            return

        if error:
            subject = f"[ARCA] Error al importar Libro de Compras - {company.name}"
            body = f"<p>Ocurrió un error al procesar <b>{filename}</b> para <b>{company.name}</b>:</p><p>{error}</p>"
        else:
            subject = (
                f"[ARCA] Importación Libro de Compras - {company.name} "
                f"({summary['imported_count']} nuevas, {summary['failed_count']} con error)"
            )
            body = f"""
                <p>Resultado de la importación automática de <b>{filename}</b> para <b>{company.name}</b>:</p>
                <ul>
                    <li>Comprobantes en el archivo: {summary['total_lines']}</li>
                    <li>Importados como nuevos: {summary['imported_count']}</li>
                    <li>Ya existentes (omitidos): {summary['duplicates_count']}</li>
                    <li>Con error (requieren revisión manual en el asistente): {summary['failed_count']}</li>
                </ul>
            """
            if summary['errors']:
                rows = "".join(
                    f"<tr><td>{e['partner_name']}</td><td>{e['document']}</td>"
                    f"<td>{e['amount_total']}</td><td>{e['error']}</td></tr>"
                    for e in summary['errors']
                )
                body += f"""
                    <table border="1" cellpadding="4" style="border-collapse: collapse;">
                        <tr><th>Proveedor</th><th>Comprobante</th><th>Importe</th><th>Error</th></tr>
                        {rows}
                    </table>
                """

        try:
            env['mail.mail'].create({
                'subject': subject,
                'body_html': body,
                'email_to': to_email,
            }).send()
        except Exception:
            # Un servidor sin remitente configurado (típico en Staging) no
            # debe hacer fallar la importación en sí, que ya se completó.
            _logger.exception("No se pudo enviar el mail de notificación de importación ARCA")
