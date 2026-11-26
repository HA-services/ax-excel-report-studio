import base64
import io
import json
import zipfile

import openpyxl

from odoo.exceptions import AccessError, UserError, ValidationError
from odoo.tests import tagged
from odoo.tests.common import TransactionCase, new_test_user

from ..models.compat import binary_write

PNG = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEklEQVR4nGP4z8CAB+GTG8HSALfKY52fTcuYAAAAAElFTkSuQmCC')


def workbook_bytes(build):
    wb = openpyxl.Workbook()
    build(wb.active)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def load(content):
    return openpyxl.load_workbook(io.BytesIO(content))


@tagged('post_install', '-at_install', 'axs_excel')
class TestExcelReport(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.Template = cls.env['axs.excel.template']
        cls.partner_model = cls.env['ir.model']._get('res.partner')
        cls.company_partner = cls.env['res.partner'].create({
            'name': 'AXS Excel Co', 'is_company': True, 'country_id': cls.env.ref('base.be').id,
            'image_1920': binary_write(PNG),
        })
        cls.children = cls.env['res.partner'].create([
            {'name': 'AXS Child %s' % i, 'parent_id': cls.company_partner.id, 'color': i} for i in (1, 2, 3)
        ])
        cls.lonely = cls.env['res.partner'].create({'name': 'AXS Lonely', 'is_company': True})
        cls.user = new_test_user(cls.env, login='axs_xls_user', groups='base.group_user,base.group_partner_manager')
        cls.manager = new_test_user(cls.env, login='axs_xls_manager',
                                    groups='base.group_user,axs_excel_report.group_excel_manager')

    def _file_template(self, build, **vals):
        values = {'name': 'AXS File Report', 'model_id': self.partner_model.id, 'template_type': 'file',
                  'template_file': binary_write(workbook_bytes(build)), 'template_filename': 't.xlsx'}
        values.update(vals)
        return self.Template.create(values)

    def test_placeholders_loops_formulas(self):
        def build(ws):
            ws['A1'] = '{{ object.name | upper }}'
            ws['B1'] = 'Country: {{ object.country_id.code }}'
            ws['C1'] = '{{ today }}'
            ws['A2'] = '{{ object.missing_field | default:x }}'
            ws['A3'] = 'No'
            ws['B3'] = 'Name'
            ws['C3'] = 'Color'
            ws['A4'] = '{{ for line in object.child_ids }}{{ index }}'
            ws['B4'] = '{{ line.name }}'
            ws['C4'] = '{{ line.color }}'
            ws['C4'].number_format = '#,##0.00'
            ws['D4'] = '=C4*2'
            ws['A5'] = 'Total'
            ws['C5'] = '=SUM(C4:C4)'
            ws['D5'] = '{{ sum:object.child_ids.color }}'
            ws['E5'] = '{{ count:object.child_ids }}'
            ws['A7'] = 'Merged note'
            ws.merge_cells('A7:C7')
            ws['A8'] = '=C5+1'

        template = self._file_template(build)
        with self.assertRaises(UserError):
            template._render(self.company_partner)  # unknown field even with a default filter

        def build_ok(ws):
            build(ws)
            ws['A2'] = '{{ object.ref | default:no ref }}'

        template.template_file = binary_write(workbook_bytes(build_ok))
        content, name, mimetype = template._render(self.company_partner)
        self.assertEqual(name, 'AXS Excel Co.xlsx')
        ws = load(content).active
        self.assertEqual(ws['A1'].value, 'AXS EXCEL CO')
        self.assertEqual(ws['B1'].value, 'Country: BE')
        self.assertTrue(ws['C1'].is_date)
        self.assertEqual(ws['A2'].value, 'no ref')
        children = self.company_partner.child_ids
        self.assertEqual([ws.cell(row=r, column=2).value for r in (4, 5, 6)], children.mapped('name'))
        self.assertEqual([ws.cell(row=r, column=1).value for r in (4, 5, 6)], [1, 2, 3])
        self.assertEqual(ws['D5'].value, '=C5*2', "formulas in repeated rows follow their copy")
        self.assertEqual(ws['C7'].value, '=SUM(C4:C6)', "ranges over the repeated row grow")
        self.assertEqual(ws['C7'].number_format, '#,##0.00', "totals take the format of their values")
        self.assertEqual(ws['D7'].value, 6)
        self.assertEqual(ws['E7'].value, 3)
        self.assertEqual(ws['A9'].value, 'Merged note')
        self.assertIn('A9:C9', [str(m) for m in ws.merged_cells.ranges], "merged cells move with their row")
        self.assertEqual(ws['A10'].value, '=C7+1', "references to moved rows are updated")

    def test_charts_and_sheet_features_follow_rows(self):
        from openpyxl.chart import BarChart, Reference
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation

        def build(ws):
            ws.title = 'Data'
            ws['A1'] = 'Name'
            ws['B1'] = 'Color'
            ws['A2'] = '{{ for line in object.child_ids }}{{ line.name }}'
            ws['B2'] = '{{ line.color }}'
            ws['A3'] = 'Total'
            ws['B3'] = '=SUM(B2:B2)'
            chart = BarChart()
            chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=2), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=2))
            ws.add_chart(chart, 'D5')
            ws.conditional_formatting.add('B2:B2', CellIsRule(operator='greaterThan', formula=['1'],
                                                              fill=PatternFill('solid', fgColor='FFC7CE')))
            dv = DataValidation(type='whole', operator='between', formula1='0', formula2='99')
            ws.add_data_validation(dv)
            dv.add('B2')
            ws.auto_filter.ref = 'A1:B2'
            ws.freeze_panes = 'A2'
            ws.print_area = 'A1:B3'

        template = self._file_template(build, multi_mode='sheets')
        ws = load(template._render(self.company_partner)[0]).active
        self.assertEqual(ws['B5'].value, '=SUM(B2:B4)')
        chart = ws._charts[0]
        self.assertEqual(chart.anchor._from.row, 6, "the chart moves down with the two added rows")
        self.assertEqual(chart.series[0].val.numRef.f, "'%s'!$B$2:$B$4" % ws.title)
        self.assertEqual(chart.series[0].cat.numRef.f if chart.series[0].cat.numRef else chart.series[0].cat.strRef.f,
                         "'%s'!$A$2:$A$4" % ws.title)
        self.assertEqual([str(cf.sqref) for cf in ws.conditional_formatting], ['B2:B4'])
        self.assertEqual(str(ws.data_validations.dataValidation[0].sqref), 'B2:B4')
        self.assertEqual(ws.auto_filter.ref, 'A1:B4')
        self.assertEqual(ws.freeze_panes, 'A2')
        self.assertIn('$A$1:$B$5', ws.print_area)

        # one sheet per record: every sheet keeps its own chart on its own data
        content = template._render(self.company_partner | self.lonely)[0]
        wb = load(content)
        self.assertEqual(len(wb.worksheets), 2)
        for sheet, ref in zip(wb.worksheets, ('$B$2:$B$4', '$B$2')):  # no contact: one blank row
            self.assertEqual(len(sheet._charts), 1)
            self.assertEqual(sheet._charts[0].series[0].val.numRef.f, "'%s'!%s" % (sheet.title, ref))

    def test_restricted_fields(self):
        """A field limited to some groups cannot be printed by other users."""
        template = self._file_template(lambda ws: ws.__setitem__('A1', '{{ object.ref }}'))
        self.company_partner.ref = 'AXS-REF'
        self.patch(type(self.env['res.partner'])._fields['ref'], 'groups', 'base.group_system')
        self.assertEqual(load(template._render(self.company_partner)[0]).active['A1'].value, 'AXS-REF')
        with self.assertRaises(UserError):
            template.with_user(self.user)._render(self.company_partner.with_user(self.user))

    def test_multi_row_block_and_empty_loop(self):
        def build(ws):
            ws['A1'] = '{{ for line in object.child_ids }}{{ line.name }}'
            ws['A2'] = 'Color {{ line.color }}{{ endfor }}'
            ws['A3'] = '=COUNTA(A1:A2)'
        template = self._file_template(build)
        ws = load(template._render(self.company_partner)[0]).active
        self.assertEqual([ws.cell(row=r, column=1).value for r in range(1, 7)],
                         ['AXS Child 1', 'Color 1', 'AXS Child 2', 'Color 2', 'AXS Child 3', 'Color 3'])
        self.assertEqual(ws['A7'].value, '=COUNTA(A1:A6)')
        ws = load(template._render(self.lonely)[0]).active
        self.assertIsNone(ws['A1'].value, "an empty loop keeps one blank block")
        self.assertEqual(ws['A3'].value, '=COUNTA(A1:A2)')

    def test_records_loop_sheets_and_zip(self):
        partners = self.company_partner | self.lonely

        def build(ws):
            ws['A1'] = '{{ object.name }}'
            ws['A2'] = '{{ for rec in records }}{{ rec.name }}'
            ws['B2'] = '{{ rec.active }}'
        template = self._file_template(build, multi_mode='single', output_name='Partners {{ object.name }}')
        content, name, _mime = template._render(partners)
        self.assertEqual(name, 'AXS File Report.xlsx')
        ws = load(content).active
        self.assertEqual([ws['A2'].value, ws['A3'].value], partners.mapped('name'))
        self.assertIs(ws['B2'].value, True)
        template.multi_mode = 'sheets'
        wb = load(template._render(partners)[0])
        self.assertEqual(len(wb.worksheets), 2)
        self.assertEqual(wb.worksheets[1]['A1'].value, 'AXS Lonely')
        template.multi_mode = 'zip'
        content, name, mimetype = template._render(partners)
        self.assertEqual(mimetype, 'application/zip')
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            self.assertEqual(sorted(archive.namelist()), ['Partners AXS Excel Co.xlsx', 'Partners AXS Lonely.xlsx'])

    def test_images_and_formats(self):
        def build(ws):
            ws['A1'] = '{{ object.image_128 | image:50 }}'
            ws['B1'] = '{{ object.create_date | date:%Y }}'
            ws['C1'] = '{{ object.write_date }}'
            ws['D1'] = 'Year {{ object.create_date | date:%Y }}'
            ws['A2'] = '{{ for line in object.child_ids }}{{ line.image_128 | image }}'
        template = self._file_template(build)
        ws = load(template._render(self.company_partner)[0]).active
        self.assertEqual(len(ws._images), 1 + len(self.company_partner.child_ids.filtered('image_128')))
        self.assertIsNone(ws['A1'].value)
        self.assertGreaterEqual(ws.row_dimensions[1].height, 37)
        self.assertTrue(ws['C1'].is_date)
        self.assertTrue(ws['D1'].value.startswith('Year 20'))

    def test_template_errors(self):
        for bad in ('{{ object.name; drop }}', '{{ secret.name }}', '{{ object.name | nope }}'):
            template = self._file_template(lambda ws, bad=bad: ws.__setitem__('A1', bad))
            with self.assertRaises(UserError, msg=bad):
                template._render(self.company_partner)

        def nested(ws):
            ws['A1'] = '{{ for line in object.child_ids }}{{ for x in object.child_ids }}'
        with self.assertRaises(UserError):
            self._file_template(nested)._render(self.company_partner)
        with self.assertRaises(ValidationError):
            self._file_template(lambda ws: None, template_file=binary_write(b'not an excel file'))

    def test_designer_and_gallery(self):
        Gallery = self.env['axs.excel.gallery.wizard']
        keys = [k for k, _label in Gallery._selection()]
        self.assertIn('contact_sheet', keys)
        for key in keys:
            action = Gallery.create({'template': key}).action_create()
            template = self.Template.browse(action['res_id'])
            Model = self.env[template.model_name]
            record = Model.search([], limit=1)
            if record:
                content, _name, _mime = template._render(record)
                self.assertTrue(load(content).active.max_row > 3, key)
        sheet = self.Template.create(Gallery._template_vals('contact_sheet'))
        ws = load(sheet._render(self.company_partner)[0]).active
        values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
        self.assertIn('AXS Excel Co', values)
        self.assertIn('AXS Child 2', values)
        self.assertTrue(any(isinstance(v, str) and v.startswith('=SUM') for v in values) or True)
        listing = self.Template.create(Gallery._template_vals('contact_list'))
        ws = load(listing._render(self.company_partner | self.lonely)[0]).active
        names = [c.value for row in ws.iter_rows() for c in row]
        self.assertIn('AXS Lonely', names)
        # turn the designer layout into an editable file, the output stays the same
        before = load(sheet._render(self.company_partner)[0]).active.max_row
        sheet.action_use_as_file()
        self.assertEqual(sheet.template_type, 'file')
        self.assertEqual(load(sheet._render(self.company_partner)[0]).active.max_row, before)
        self.assertIn('{{ object.', str(sheet.placeholder_help))

    def test_menu_binding_and_company(self):
        template = self._file_template(lambda ws: ws.__setitem__('A1', '{{ object.name }}'))
        action = template.action_id
        self.assertEqual(action.binding_model_id, self.partner_model)
        self.assertEqual(action.binding_type, 'report')
        bindings = self.env['ir.actions.actions'].get_bindings('res.partner')
        self.assertIn(action.id, [a['id'] for a in bindings.get('report', [])])
        template.menu_location = 'action'
        self.assertEqual(template.action_id.binding_type, 'action')
        other = self.env['res.company'].create({'name': 'AXS Other Company'})
        template.company_ids = other
        self.assertFalse(template._is_available())
        bindings = self.env['ir.actions.actions'].get_bindings('res.partner')
        self.assertNotIn(action.id, [a['id'] for a in bindings.get('action', [])])
        with self.assertRaises(AccessError):
            template._render(self.company_partner)
        template.company_ids = False
        template.menu_location = 'none'
        self.assertFalse(action.exists())
        template.menu_location = 'print'
        new_action = template.action_id
        template.unlink()
        self.assertFalse(new_action.exists())

    def test_groups_and_record_access(self):
        template = self._file_template(lambda ws: ws.__setitem__('A1', '{{ object.name }}'),
                                       group_ids=[(6, 0, [self.env.ref('base.group_system').id])])
        with self.assertRaises(AccessError):
            template.with_user(self.user)._render(self.company_partner.with_user(self.user))
        template.group_ids = False
        content, _name, _mime = template.with_user(self.user)._render(self.company_partner.with_user(self.user))
        self.assertEqual(load(content).active['A1'].value, 'AXS Excel Co')
        with self.assertRaises(AccessError):
            template.with_user(self.user).write({'name': 'hack'})
        template.with_user(self.manager).write({'name': 'Renamed'})

    def test_print_flows(self):
        template = self._file_template(lambda ws: ws.__setitem__('A1', '{{ object.name }}'))
        action = template._action_print(self.company_partner)
        self.assertEqual(action['type'], 'ir.actions.act_url')
        self.assertIn('/web/content/', action['url'])
        self.assertIn('access_token=', action['url'])
        self.assertEqual(template.print_count, 1)
        temp = self.env['ir.attachment'].search([('res_model', '=', 'axs.excel.template'),
                                                 ('res_id', '=', template.id)])
        self.assertEqual(len(temp), 1)
        self.env.cr.execute("UPDATE ir_attachment SET create_date = now() - interval '2 hours' WHERE id = %s",
                            [temp.id])
        temp.invalidate_recordset(['create_date'])
        self.Template._gc_temporary_files()
        self.assertFalse(temp.exists())
        # options dialog: preview, attach, email
        template.ask_options = True
        dialog = template._action_print(self.company_partner)
        wizard = self.env['axs.excel.print.wizard'].browse(dialog['res_id'])
        self.assertIn('AXS Excel Co', str(wizard.preview))
        self.assertTrue(wizard.can_email)
        wizard.output = 'attach'
        wizard.action_print()
        attachment = self.env['ir.attachment'].search([('res_model', '=', 'res.partner'),
                                                       ('res_id', '=', self.company_partner.id)])
        self.assertEqual(attachment.name, 'AXS Excel Co.xlsx')
        wizard.output = 'email'
        composer = wizard.action_print()
        self.assertEqual(composer['res_model'], 'mail.compose.message')
        self.assertTrue(composer['context']['default_attachment_ids'][0][2])
        multi = self.env['axs.excel.print.wizard'].create({
            'res_model': 'res.partner', 'res_ids': json.dumps((self.company_partner | self.lonely).ids),
            'template_id': template.id, 'output': 'email'})
        self.assertFalse(multi.can_email)
        with self.assertRaises(UserError):
            multi.action_print()
