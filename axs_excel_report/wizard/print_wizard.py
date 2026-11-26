import json

from odoo import _, api, fields, models
from odoo.exceptions import UserError

from ..models.excel_engine import TemplateError, preview_html


class AxsExcelPrintWizard(models.TransientModel):
    _name = 'axs.excel.print.wizard'
    _description = 'Print an Excel Report'

    res_model = fields.Char(required=True)
    res_ids = fields.Char(required=True, help="JSON list of the record ids")
    record_count = fields.Integer(compute='_compute_record_count')
    template_id = fields.Many2one('axs.excel.template', 'Report', required=True,
                                  domain="[('model_name', '=', res_model)]")
    output = fields.Selection([
        ('download', 'Download'),
        ('attach', 'Attach to the record'),
        ('email', 'Send by email'),
    ], default='download', required=True)
    preview = fields.Html(compute='_compute_preview', sanitize=False)
    can_email = fields.Boolean(compute='_compute_record_count')

    @api.depends('res_ids', 'res_model')
    def _compute_record_count(self):
        for wizard in self:
            ids = json.loads(wizard.res_ids or '[]')
            wizard.record_count = len(ids)
            model = self.env[wizard.res_model] if wizard.res_model in self.env else None
            wizard.can_email = bool(model is not None and len(ids) == 1 and hasattr(model, 'message_post'))

    def _records(self):
        self.ensure_one()
        if self.res_model not in self.env:
            raise UserError(_("The model %s is not available.", self.res_model))
        return self.env[self.res_model].browse(json.loads(self.res_ids or '[]')).exists()

    @api.depends('template_id', 'res_ids')
    def _compute_preview(self):
        for wizard in self:
            if not wizard.template_id or not wizard.res_ids:
                wizard.preview = False
                continue
            try:
                records = wizard._records()[:1]
                content, _name, _mime = wizard.template_id.with_context(
                    axs_excel_preview=True)._render(records) if records else (None, None, None)
                wizard.preview = preview_html(content, texts={
                    'more': _("%s more rows in the file."),
                    'pictures': _("%s picture(s) in the sheet."),
                    'formulas': _("Formulas, shown in grey, are calculated when the file opens."),
                }) if content else False
            except (UserError, TemplateError) as e:
                wizard.preview = '<div class="alert alert-warning mb-0">%s</div>' % (
                    str(e.args[0] if e.args else e).replace('<', '&lt;'))

    @api.model
    def _open_action(self, records):
        templates = self.env['axs.excel.template']._available_for(records._name)
        default = self.env.context.get('default_template_id') or (templates[:1].id if templates else False)
        wizard = self.create({
            'res_model': records._name,
            'res_ids': json.dumps(records.ids),
            'template_id': default,
        })
        return {
            'type': 'ir.actions.act_window',
            'name': _("Excel Report"),
            'res_model': self._name,
            'res_id': wizard.id,
            'view_mode': 'form',
            'views': [(False, 'form')],
            'target': 'new',
        }

    def action_print(self):
        self.ensure_one()
        template = self.template_id
        records = self._records()
        if not records:
            raise UserError(_("The selected records do not exist anymore."))
        if self.output == 'download':
            content, name, mimetype = template._render(records)
            template.sudo().print_count += 1
            return template._download_action(template._make_attachment(records, content, name, mimetype))
        if self.output == 'attach':
            for record in records:
                content, name, mimetype = template._render(record)
                template._make_attachment(record, content, name, mimetype, attach=True)
            template.sudo().print_count += len(records)
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'type': 'success',
                    'message': _("%s Excel file(s) attached.", len(records)),
                    'next': {'type': 'ir.actions.act_window_close'},
                },
            }
        if not self.can_email:
            raise UserError(_("Email is available for one record of a model with a chatter."))
        content, name, mimetype = template._render(records)
        attachment = template._make_attachment(records, content, name, mimetype, attach=True)
        template.sudo().print_count += 1
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'mail.compose.message',
            'view_mode': 'form',
            'views': [(False, 'form')],
            'target': 'new',
            'context': {
                'default_model': records._name,
                'default_res_ids': records.ids,
                'default_res_id': records.id,
                'default_composition_mode': 'comment',
                'default_subject': name.rsplit('.', 1)[0],
                'default_attachment_ids': [(6, 0, attachment.ids)],
            },
        }
