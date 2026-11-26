import io
import logging
import re
import zipfile

from markupsafe import Markup, escape

from odoo import _, api, fields, models
from odoo.exceptions import AccessError, UserError, ValidationError

from .compat import binary_read, binary_write, can_read, check_read
from .excel_engine import Renderer, TemplateError, openpyxl

_logger = logging.getLogger(__name__)

XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
FIELD_TYPES_SKIPPED = ('one2many', 'many2many', 'binary', 'properties', 'properties_definition', 'json')
MAX_RECORDS = 5000


class AxsExcelTemplate(models.Model):
    _name = 'axs.excel.template'
    _description = 'Excel Report Template'
    _order = 'model_id, sequence, name, id'

    name = fields.Char('Report Name', required=True, translate=True)
    sequence = fields.Integer(default=10)
    active = fields.Boolean(default=True)
    description = fields.Text(translate=True)
    model_id = fields.Many2one('ir.model', 'Model', required=True, ondelete='cascade',
                               domain="[('transient', '=', False)]")
    model_name = fields.Char(related='model_id.model', string='Model Name', store=True)
    company_ids = fields.Many2many('res.company', 'axs_excel_template_company_rel', 'template_id', 'company_id',
                                   'Companies', help="The report is offered in these companies only. "
                                                     "Leave empty for all companies.")
    group_ids = fields.Many2many('res.groups', 'axs_excel_template_group_rel', 'template_id', 'group_id',
                                 'Allowed Groups', help="Users of these groups can print the report. "
                                                        "Leave empty for every user who can read the records.")
    template_type = fields.Selection([
        ('designer', 'Designer'),
        ('file', 'Excel File'),
    ], 'Template Source', default='designer', required=True,
        help="Designer: choose the fields, the workbook is built for you.\n"
             "Excel File: upload your own workbook with placeholders.")
    template_file = fields.Binary('Excel Template', attachment=True)
    template_filename = fields.Char('File Name')
    output_name = fields.Char('Output File Name', default='{{ object.display_name }}', translate=True,
                              help="Name of the downloaded file, placeholders allowed. "
                                   "Example: Quotation {{ object.name }}")
    multi_mode = fields.Selection([
        ('sheets', 'One sheet per record'),
        ('zip', 'One file per record (zip)'),
        ('single', 'One workbook for all records (records loop)'),
    ], 'Several Records', default='sheets', required=True,
        help="What happens when several records are selected in a list.")
    menu_location = fields.Selection([
        ('print', 'Print menu'),
        ('action', 'Action menu'),
        ('none', 'No menu entry'),
    ], 'Show In', default='print', required=True)
    view_types = fields.Selection([
        ('list,form', 'Form and list views'),
        ('form', 'Form view only'),
        ('list', 'List view only'),
    ], 'Available In', default='list,form', required=True)
    ask_options = fields.Boolean('Open Options Dialog',
                                 help="Show a dialog with a preview and the download, attach and email choices "
                                      "instead of downloading the file directly.")
    action_id = fields.Many2one('ir.actions.server', 'Menu Action', readonly=True, copy=False)
    print_count = fields.Integer(readonly=True, copy=False)

    # designer
    layout = fields.Selection([
        ('record', 'Document (one record with its lines)'),
        ('list', 'Table of the selected records'),
    ], 'Layout', default='record', required=True,
        help="Document: header fields and a table of lines, one sheet per record.\n"
             "Table: one row per selected record, for list exports.")
    title = fields.Char(default='{{ object.display_name }}', translate=True)
    show_company = fields.Boolean('Company Header', default=True)
    show_logo = fields.Boolean('Company Logo', default=True)
    header_line_ids = fields.One2many('axs.excel.template.line', 'template_id', 'Fields',
                                      domain=[('section', '=', 'header')], context={'default_section': 'header'},
                                      copy=True)
    table_field_id = fields.Many2one('ir.model.fields', 'Lines',
                                     domain="[('model_id', '=', model_id), ('ttype', 'in', ('one2many', 'many2many'))]",
                                     ondelete='set null', help="The lines shown as a table, for example the order lines.")
    table_model_id = fields.Many2one('ir.model', compute='_compute_table_model', string='Lines Model')
    table_line_ids = fields.One2many('axs.excel.template.line', 'template_id', 'Columns',
                                     domain=[('section', '=', 'table')], context={'default_section': 'table'},
                                     copy=True)
    footer = fields.Text('Footer Note', translate=True)
    accent_color = fields.Char('Accent Color', default='#1b264f')

    placeholder_help = fields.Html(compute='_compute_placeholder_help', sanitize=False)

    @api.depends('table_field_id')
    def _compute_table_model(self):
        for template in self:
            relation = template.table_field_id.relation
            template.table_model_id = self.env['ir.model']._get(relation) if relation else False

    @api.depends('model_id')
    def _compute_placeholder_help(self):
        for template in self:
            template.placeholder_help = template._placeholder_help() if template.model_id else False

    @api.constrains('template_type', 'template_file', 'output_name', 'accent_color', 'model_id')
    def _check_template(self):
        for template in self:
            if template.template_type == 'file' and template.template_file:
                data = binary_read(template.with_context(bin_size=False).template_file)
                if not zipfile.is_zipfile(io.BytesIO(data)):
                    raise ValidationError(_("The template must be an Excel workbook (.xlsx)."))
                try:
                    openpyxl.load_workbook(io.BytesIO(data))
                except Exception as e:
                    raise ValidationError(_("The Excel template cannot be opened: %s", e)) from e
            if template.accent_color and not re.fullmatch(r'#[0-9a-fA-F]{6}', template.accent_color):
                raise ValidationError(_("The accent color must look like #1b264f."))
            if template.model_id and template.model_id.model not in self.env:
                raise ValidationError(_("The model %s is not available.", template.model_id.model))

    # ------------------------------------------------------------------ menu binding
    @api.model_create_multi
    def create(self, vals_list):
        templates = super().create(vals_list)
        templates._sync_action()
        return templates

    def write(self, vals):
        res = super().write(vals)
        if {'name', 'model_id', 'menu_location', 'view_types', 'group_ids', 'active', 'sequence'} & set(vals):
            self._sync_action()
        return res

    def unlink(self):
        actions = self.mapped('action_id')
        res = super().unlink()
        actions.sudo().unlink()
        return res

    def copy(self, default=None):
        default = dict(default or {})
        if len(self) == 1:
            default.setdefault('name', _("%s (copy)", self.name))
        return super().copy(default)

    def _sync_action(self):
        Action = self.env['ir.actions.server'].sudo()
        groups_field = 'group_ids' if 'group_ids' in Action._fields else 'groups_id'
        for template in self.sudo():
            if template.menu_location == 'none' or not template.active:
                if template.action_id:
                    template.action_id.unlink()
                continue
            vals = {
                'name': template.name,
                'model_id': template.model_id.id,
                'binding_model_id': template.model_id.id,
                'binding_type': 'report' if template.menu_location == 'print' else 'action',
                'binding_view_types': template.view_types,
                'state': 'code',
                'code': "action = env['axs.excel.template'].browse(%d)._action_print(records)" % template.id,
                groups_field: [(6, 0, template.group_ids.ids)],
                'sequence': template.sequence,
            }
            if template.action_id:
                template.action_id.write(vals)
            else:
                template.action_id = Action.create(vals)

    # ------------------------------------------------------------------ availability
    def _is_available(self, user=None):
        """Whether the current user may print this template in the current company."""
        self.ensure_one()
        user = user or self.env.user
        template = self.sudo()
        if not template.active:
            return False
        if template.company_ids and self.env.company not in template.company_ids:
            return False
        if template.group_ids and not (template.group_ids & user_groups(user)):
            return False
        return True

    @api.model
    def _available_for(self, model_name):
        templates = self.sudo().search([('model_name', '=', model_name)])
        return templates.filtered(lambda t: t._is_available())

    # ------------------------------------------------------------------ rendering
    def _template_bytes(self):
        self.ensure_one()
        if self.template_type == 'file':
            if not self.template_file:
                raise UserError(_("Upload an Excel template for '%s' first.", self.name))
            return binary_read(self.with_context(bin_size=False, bin_size_template_file=False).template_file)
        return self._designer_workbook()

    def _render(self, records):
        """Render the report for ``records`` with the access rights of the current user.

        :returns: ``(content, file name, mimetype)``
        """
        self.ensure_one()
        if openpyxl is None:
            raise UserError(_("The Python library openpyxl is required to build Excel reports."))
        if not records:
            raise UserError(_("Select at least one record."))
        if records._name != self.model_name:
            raise UserError(_("This report is made for %s records.", self.model_id.name))
        if not self._is_available():
            raise AccessError(_("You are not allowed to print '%s'.", self.name))
        if len(records) > MAX_RECORDS:
            raise UserError(_("Select at most %s records for one Excel report.", MAX_RECORDS))
        check_read(records)
        template = self.sudo()
        source = template._template_bytes()
        renderer_env = records.env
        try:
            if template.multi_mode == 'zip' and template.layout != 'list' and len(records) > 1:
                buffer = io.BytesIO()
                used = set()
                with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
                    for record in records:
                        content = Renderer(renderer_env, record).render_workbook(source)
                        name = self._unique_name(template._file_name(record), used)
                        archive.writestr(name, content)
                return buffer.getvalue(), '%s.zip' % self._safe(template.with_context(lang=self.env.lang).name), \
                    'application/zip'
            mode = 'sheets' if template.multi_mode == 'sheets' and template.layout != 'list' else 'single'
            content = Renderer(renderer_env, records).render_workbook(source, mode=mode)
        except TemplateError as e:
            raise UserError(_("The Excel template of '%s' has an error: %s", template.name, e)) from e
        name = template._file_name(records[:1]) if len(records) == 1 else '%s.xlsx' % self._safe(
            template.with_context(lang=self.env.lang).name)
        return content, name, XLSX_MIME

    def _file_name(self, record):
        try:
            name = Renderer(record.env, record).render_text(self.output_name or '{{ object.display_name }}',
                                                            {'object': record})
        except TemplateError:
            name = record.display_name
        return '%s.xlsx' % self._safe(name or self.name)

    @staticmethod
    def _safe(name):
        return re.sub(r'[\\/:*?"<>|\r\n]+', '-', str(name or 'report')).strip()[:120] or 'report'

    @staticmethod
    def _unique_name(name, used):
        base, ext = name.rsplit('.', 1)
        n = 2
        while name in used:
            name = '%s (%s).%s' % (base, n, ext)
            n += 1
        used.add(name)
        return name

    def _make_attachment(self, records, content, name, mimetype, attach=False):
        """Store the generated file: on the record (attach) or as a temporary download."""
        vals = {'name': name, 'raw': content, 'mimetype': mimetype}
        if attach and len(records) == 1:
            vals.update(res_model=records._name, res_id=records.id)
            return records.env['ir.attachment'].create(vals)
        vals.update(res_model=self._name, res_id=self.id, description='axs_excel_report:temporary')
        return self.env['ir.attachment'].sudo().create(vals).with_env(self.env)

    def _download_action(self, attachment):
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true&access_token=%s' % (
                attachment.id, attachment.sudo().generate_access_token()[0]),
            'target': 'self',
        }

    def _action_print(self, records):
        """Entry point of the Print / Action menu entries."""
        self.ensure_one()
        if not records:
            records = self.env[self.model_name].browse(self.env.context.get('active_ids') or [])
        if self.sudo().ask_options:
            return self.env['axs.excel.print.wizard'].with_context(
                default_template_id=self.id, active_model=records._name, active_ids=records.ids,
            )._open_action(records)
        content, name, mimetype = self._render(records)
        self.sudo().print_count += 1
        return self._download_action(self._make_attachment(records, content, name, mimetype))

    @api.model
    def _gc_temporary_files(self):
        """Cron: remove the temporary downloads older than one hour."""
        limit = fields.Datetime.subtract(fields.Datetime.now(), hours=1)
        self.env['ir.attachment'].sudo().search([
            ('res_model', '=', self._name), ('description', '=', 'axs_excel_report:temporary'),
            ('create_date', '<', limit),
        ]).unlink()

    # ------------------------------------------------------------------ designer
    def _designer_workbook(self):
        """Build the template workbook from the designer settings."""
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
        self.ensure_one()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self._safe(self.with_context(lang=self.env.lang).name)[:31] or 'Report'
        accent = (self.accent_color or '#1b264f').lstrip('#')
        light = 'F3F5F9'
        thin = Side(style='thin', color='D7DCE5')
        columns = max(4, len(self.table_line_ids))
        row = 1
        if self.show_company:
            if self.show_logo:
                ws.cell(row=row, column=columns, value='{{ company.logo | image:48 }}')
            ws.cell(row=row, column=1, value='{{ company.name }}').font = Font(bold=True, size=12, color=accent)
            ws.cell(row=row + 1, column=1, value='{{ company.partner_id.contact_address_complete | default: }}'
                    if 'contact_address_complete' in self.env['res.partner']._fields
                    else '{{ company.email }}').font = Font(size=9, color='6B7280')
            row += 3
        title = ws.cell(row=row, column=1, value=self.title or '{{ object.display_name }}')
        title.font = Font(bold=True, size=16, color=accent)
        ws.row_dimensions[row].height = 26
        row += 2
        if self.layout == 'list':
            title.value = self.title if self.title and 'object.' not in self.title else self.with_context(
                lang=self.env.lang).name
            ws.cell(row=row + 1, column=1, value=_("Records: {{ count }}, {{ today }}")).font = Font(size=9, color='6B7280')
            row += 3
            lines = list(self.header_line_ids)
            columns = max(columns, len(lines))
            for c, line in enumerate(lines, start=1):
                cell = ws.cell(row=row, column=c, value=line.name)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor=accent)
                cell.alignment = Alignment(horizontal='right' if line.is_numeric else 'left', vertical='center')
            ws.row_dimensions[row].height = 20
            row += 1
            loop_row = row
            for c, line in enumerate(lines, start=1):
                text = '{{ rec.%s }}' % line.path if not line.is_image else '{{ rec.%s | image:40 }}' % line.path
                if c == 1:
                    text = '{{ for rec in records }}' + text
                cell = ws.cell(row=row, column=c, value=text)
                cell.border = Border(bottom=thin)
                cell.alignment = Alignment(horizontal='right' if line.is_numeric else 'left', vertical='center')
            row += 1
            if any(line.total for line in lines):
                for c, line in enumerate(lines, start=1):
                    cell = ws.cell(row=row, column=c)
                    cell.fill = PatternFill('solid', fgColor=light)
                    cell.font = Font(bold=True)
                    if line.total:
                        letter = get_column_letter(c)
                        cell.value = '=SUM(%s%s:%s%s)' % (letter, loop_row, letter, loop_row)
                        cell.alignment = Alignment(horizontal='right')
                        cell.number_format = '#,##0.00'
                    elif c == 1:
                        cell.value = _("Total")
                row += 1
            ws.freeze_panes = 'A%s' % loop_row
            row += 1
            pairs = []
            table_lines = lines
        else:
            pairs = list(self.header_line_ids)
            table_lines = list(self.table_line_ids)
        # header fields, two per row
        for i in range(0, len(pairs), 2):
            for j, line in enumerate(pairs[i:i + 2]):
                col = 1 + j * max(2, columns // 2)
                label = ws.cell(row=row, column=col, value=line.name)
                label.font = Font(bold=True, color='6B7280', size=9)
                value = ws.cell(row=row, column=col + 1, value='{{ object.%s }}' % line.path)
                value.alignment = Alignment(horizontal='left')
            row += 1
        if pairs:
            row += 1
        # table
        if self.layout == 'record' and self.table_field_id and self.table_line_ids:
            for c, line in enumerate(self.table_line_ids, start=1):
                cell = ws.cell(row=row, column=c, value=line.name)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor=accent)
                cell.alignment = Alignment(horizontal='right' if line.is_numeric else 'left', vertical='center')
            ws.row_dimensions[row].height = 20
            row += 1
            loop_row = row
            for c, line in enumerate(self.table_line_ids, start=1):
                text = '{{ line.%s }}' % line.path
                if line.is_image:
                    text = '{{ line.%s | image:40 }}' % line.path
                if c == 1:
                    text = '{{ for line in object.%s }}%s' % (self.table_field_id.name, text)
                cell = ws.cell(row=row, column=c, value=text)
                cell.border = Border(bottom=thin)
                cell.alignment = Alignment(horizontal='right' if line.is_numeric else 'left', vertical='center')
            row += 1
            if any(line.total for line in self.table_line_ids):
                for c, line in enumerate(self.table_line_ids, start=1):
                    cell = ws.cell(row=row, column=c)
                    cell.fill = PatternFill('solid', fgColor=light)
                    cell.font = Font(bold=True)
                    if line.total:
                        letter = get_column_letter(c)
                        cell.value = '=SUM(%s%s:%s%s)' % (letter, loop_row, letter, loop_row)
                        cell.alignment = Alignment(horizontal='right')
                        cell.number_format = '#,##0.00'
                    elif c == 1:
                        cell.value = _("Total")
                row += 1
            row += 1
        if self.footer:
            ws.cell(row=row, column=1, value=self.footer).font = Font(italic=True, color='6B7280', size=9)
            row += 1
        ws.cell(row=row + 1, column=1, value=_("Printed on {{ now }} by {{ user.name }}")).font = Font(
            size=8, color='9CA3AF')
        widths = [max(14, min(40, len(line.name) + 6)) for line in table_lines] or [22] * columns
        for c in range(1, columns + 1):
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1] if c - 1 < len(widths) else 22
        if table_lines:
            ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 34)
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = 'landscape' if columns > 6 else 'portrait'
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def action_download_template(self):
        """Download the template workbook (the designer layout or the uploaded file) to customize it."""
        self.ensure_one()
        content = self._template_bytes()
        name = self.template_filename if self.template_type == 'file' and self.template_filename else \
            '%s - template.xlsx' % self._safe(self.name)
        return self._download_action(self._make_attachment(self.env[self.model_name], content, name, XLSX_MIME))

    def action_use_as_file(self):
        """Turn the designer layout into an uploaded template that can be edited in Excel."""
        for template in self:
            content = template._designer_workbook()
            template.write({
                'template_file': binary_write(content),
                'template_filename': '%s.xlsx' % self._safe(template.name),
                'template_type': 'file',
            })
        return True

    def action_test_print(self):
        """Print the report for the most recent record of the model."""
        self.ensure_one()
        Model = self.env[self.model_name]
        if not can_read(Model):
            raise AccessError(_("You cannot read %s records.", self.model_id.name))
        record = Model.search([], limit=1, order='id desc')
        if not record:
            raise UserError(_("There is no %s record to print yet.", self.model_id.name))
        return self.env['axs.excel.print.wizard'].with_context(default_template_id=self.id)._open_action(record)

    # ------------------------------------------------------------------ placeholder help
    def _placeholder_help(self):
        self.ensure_one()
        Model = self.env[self.model_name] if self.model_name in self.env else None
        if Model is None:
            return False
        rows = []
        loops = []
        for name, field in sorted(Model._fields.items(), key=lambda kv: kv[1].string or kv[0]):
            if name in ('id', '__last_update') or not field.store and not field.related:
                continue
            if field.type in ('one2many', 'many2many'):
                loops.append((field, name))
                continue
            if field.type in FIELD_TYPES_SKIPPED and field.type != 'binary':
                continue
            code = '{{ object.%s }}' % name
            if field.type == 'binary':
                code = '{{ object.%s | image:80 }}' % name
            elif field.type == 'many2one':
                code = '{{ object.%s.name }}' % name if 'name' in self.env[field.comodel_name]._fields else code
            rows.append('<tr><td>%s</td><td><code>%s</code></td><td class="text-muted">%s</td></tr>' % (
                escape(field.string), escape(code), escape(field.type)))
        loop_rows = ''.join(
            '<tr><td>%s</td><td><code>%s</code></td><td class="text-muted">%s</td></tr>' % (
                escape(field.string), escape('{{ for line in object.%s }}' % name), escape(field.comodel_name))
            for field, name in loops)
        general = ''.join('<tr><td>%s</td><td><code>%s</code></td></tr>' % (escape(a), escape(b)) for a, b in [
            (_("Company name and logo"), '{{ company.name }}   {{ company.logo | image:60 }}'),
            (_("Current user"), '{{ user.name }}'),
            (_("Print date"), '{{ today }}   {{ now }}   {{ today | date:%d/%m/%Y }}'),
            (_("Line number in a loop"), '{{ index }}'),
            (_("Totals"), '{{ sum:object.order_line.price_subtotal }}   (sum, avg, min, max, count)'),
            (_("All printed records"), '{{ for rec in records }} ... {{ rec.name }}'),
            (_("Several rows per line"), '{{ for line in object.order_line }} ... {{ endfor }}'),
            (_("Formatting"), '| upper   | lower   | title   | round:2   | default:-   | date:%d.%m.%Y'),
        ])
        return Markup(
            '<div class="o_axs_placeholders"><h5>%s</h5><table class="table table-sm"><tbody>%s</tbody></table>'
            '<h5>%s</h5><table class="table table-sm"><tbody>%s</tbody></table>'
            '<h5>%s</h5><table class="table table-sm"><tbody>%s</tbody></table></div>'
        ) % (_("General"), Markup(general), _("Lines (loops)"), Markup(loop_rows or '<tr><td>-</td></tr>'),
             _("Fields of %s", self.model_id.name), Markup(''.join(rows)))


class AxsExcelTemplateLine(models.Model):
    _name = 'axs.excel.template.line'
    _description = 'Excel Report Field'
    _order = 'sequence, id'

    template_id = fields.Many2one('axs.excel.template', required=True, ondelete='cascade', index=True)
    section = fields.Selection([('header', 'Header'), ('table', 'Table')], required=True, default='header')
    sequence = fields.Integer(default=10)
    name = fields.Char('Label', required=True, translate=True)
    model_id = fields.Many2one('ir.model', compute='_compute_line_model', string='Line Model')
    field_id = fields.Many2one('ir.model.fields', 'Field', required=True, ondelete='cascade',
                               domain="[('model_id', '=', model_id), ('ttype', 'not in', "
                                      "('one2many', 'properties', 'properties_definition', 'json'))]")
    sub_path = fields.Char('Then', help="For a relation, the field of the related record, for example "
                                        "'email' or 'country_id.code'. Leave empty to show its name.")
    path = fields.Char(compute='_compute_path', store=True)
    total = fields.Boolean(help="Add an Excel SUM below the column.")
    is_numeric = fields.Boolean(compute='_compute_path', store=True)
    is_image = fields.Boolean(compute='_compute_path', store=True)

    @api.depends('section', 'template_id.model_id', 'template_id.table_field_id')
    def _compute_line_model(self):
        for line in self:
            if line.section == 'table':
                relation = line.template_id.table_field_id.relation
                line.model_id = self.env['ir.model']._get(relation) if relation else False
            else:
                line.model_id = line.template_id.model_id

    @api.depends('field_id', 'sub_path')
    def _compute_path(self):
        for line in self:
            path = line.field_id.name or ''
            if line.sub_path:
                path = '%s.%s' % (path, line.sub_path.strip('. '))
            line.path = path
            ttype = line.field_id.ttype
            line.is_numeric = ttype in ('integer', 'float', 'monetary') and not line.sub_path
            line.is_image = ttype == 'binary' and not line.sub_path

    @api.onchange('field_id')
    def _onchange_field_id(self):
        if self.field_id and not self.name:
            self.name = self.field_id.field_description

    @api.constrains('field_id', 'sub_path', 'total')
    def _check_path(self):
        for line in self:
            if line.sub_path and not re.fullmatch(r'[a-z_][a-z0-9_]*(\.[a-z_][a-z0-9_]*)*', line.sub_path.strip('. ')):
                raise ValidationError(_("'%s' is not a valid field path.", line.sub_path))
            if line.sub_path:
                Model = self.env[line.field_id.relation] if line.field_id.relation in self.env else None
                if Model is None:
                    raise ValidationError(_("'Then' can only follow a relation field (%s).", line.name))
                for name in line.sub_path.strip('. ').split('.'):
                    if Model is None or name not in Model._fields:
                        raise ValidationError(_("The field '%s' does not exist (column %s).", name, line.name))
                    field = Model._fields[name]
                    Model = self.env[field.comodel_name] if field.relational else None
            if line.total and not line.is_numeric:
                raise ValidationError(_("Totals are only possible on number columns (%s).", line.name))


def user_groups(user):
    """All the groups of a user, implied ones included (field renamed in Odoo 19)."""
    if 'all_group_ids' in user._fields:
        return user.sudo().all_group_ids
    return user.sudo().groups_id
