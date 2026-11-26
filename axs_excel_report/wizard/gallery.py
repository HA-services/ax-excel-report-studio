"""Ready-made Excel reports for the installed apps."""
from odoo import _ as _odoo_t
from odoo import api, fields, models
from odoo.exceptions import UserError


def _gallery(env):
    _ = getattr(env, '_', None) or _odoo_t  # noqa: F811 (env._ on Odoo 18+)
    return {
        'sale_quotation': {
            'name': _("Quotation (Excel)"), 'model': 'sale.order', 'output': 'Quotation {{ object.name }}',
            'header': [('partner_id', _("Customer")), ('date_order', _("Order Date")), ('user_id', _("Salesperson")),
                       ('validity_date', _("Expiration")), ('payment_term_id', _("Payment Terms")),
                       ('client_order_ref', _("Customer Reference"))],
            'table': 'order_line',
            'columns': [('product_id', _("Product"), False), ('name', _("Description"), False),
                        ('product_uom_qty', _("Quantity"), False), ('price_unit', _("Unit Price"), False),
                        ('price_subtotal', _("Subtotal"), True)],
            'footer': _("Prices are expressed without taxes."),
        },
        'sale_list': {
            'name': _("Sales Orders List (Excel)"), 'model': 'sale.order', 'layout': 'list',
            'output': 'Sales orders', 'title': _("Sales Orders"),
            'header': [('name', _("Order")), ('partner_id', _("Customer")), ('date_order', _("Date")),
                       ('user_id', _("Salesperson")), ('state', _("Status")),
                       ('amount_untaxed', _("Untaxed"), True), ('amount_total', _("Total"), True)],
        },
        'purchase_order': {
            'name': _("Purchase Order (Excel)"), 'model': 'purchase.order', 'output': 'Purchase {{ object.name }}',
            'header': [('partner_id', _("Vendor")), ('partner_ref', _("Vendor Reference")),
                       ('date_order', _("Order Deadline")), ('user_id', _("Buyer"))],
            'table': 'order_line',
            'columns': [('product_id', _("Product"), False), ('name', _("Description"), False),
                        ('product_qty', _("Quantity"), False), ('price_unit', _("Unit Price"), False),
                        ('price_subtotal', _("Subtotal"), True)],
        },
        'invoice': {
            'name': _("Invoice (Excel)"), 'model': 'account.move', 'output': '{{ object.name }}',
            'header': [('partner_id', _("Customer")), ('invoice_date', _("Invoice Date")),
                       ('invoice_date_due', _("Due Date")), ('payment_reference', _("Payment Reference")),
                       ('currency_id', _("Currency"))],
            'table': 'invoice_line_ids',
            'columns': [('product_id', _("Product"), False), ('name', _("Label"), False),
                        ('quantity', _("Quantity"), False), ('price_unit', _("Unit Price"), False),
                        ('price_subtotal', _("Subtotal"), True)],
        },
        'delivery': {
            'name': _("Transfer Sheet (Excel)"), 'model': 'stock.picking', 'output': '{{ object.name }}',
            'header': [('partner_id', _("Contact")), ('picking_type_id', _("Operation")),
                       ('scheduled_date', _("Scheduled Date")), ('origin', _("Source Document"))],
            'table': 'move_ids',
            'columns': [('product_id', _("Product"), False), ('product_uom_qty', _("Demand"), True),
                        ('quantity', _("Quantity"), True), ('quantity_done', _("Done"), True),
                        ('product_uom', _("Unit"), False)],
        },
        'contact_sheet': {
            'name': _("Contact Sheet (Excel)"), 'model': 'res.partner', 'output': '{{ object.name }}',
            'header': [('name', _("Name")), ('email', _("Email")), ('phone', _("Phone")), ('street', _("Street")),
                       ('city', _("City")), ('country_id', _("Country")), ('vat', _("Tax ID")),
                       ('website', _("Website")), ('image_128', _("Picture"))],
            'table': 'child_ids',
            'columns': [('name', _("Contact"), False), ('function', _("Job Position"), False),
                        ('email', _("Email"), False), ('phone', _("Phone"), False)],
        },
        'contact_list': {
            'name': _("Contact List (Excel)"), 'model': 'res.partner', 'layout': 'list', 'output': 'Contacts',
            'title': _("Contacts"),
            'header': [('name', _("Name")), ('email', _("Email")), ('phone', _("Phone")), ('city', _("City")),
                       ('country_id', _("Country")), ('user_id', _("Salesperson"))],
        },
        'product_list': {
            'name': _("Product Price List (Excel)"), 'model': 'product.template', 'layout': 'list',
            'output': 'Price list', 'title': _("Price List"),
            'header': [('image_128', _("Picture")), ('name', _("Product")), ('default_code', _("Reference")),
                       ('categ_id', _("Category")), ('list_price', _("Sales Price")),
                       ('standard_price', _("Cost"))],
        },
    }


class AxsExcelGalleryWizard(models.TransientModel):
    _name = 'axs.excel.gallery.wizard'
    _description = 'Create an Excel Report from the Gallery'

    @api.model
    def _selection(self):
        return [(key, spec['name']) for key, spec in _gallery(self.env).items() if spec['model'] in self.env]

    template = fields.Selection('_selection', 'Report', required=True)

    @api.model
    def _template_vals(self, key):
        spec = _gallery(self.env)[key]
        model = self.env['ir.model']._get(spec['model'])
        Fields = self.env['ir.model.fields']

        def lines(model_name, entries, section):
            result = []
            for seq, entry in enumerate(entries):
                name, label = entry[0], entry[1]
                total = entry[2] if len(entry) > 2 else False
                field = Fields._get(model_name, name)
                if not field or field.ttype in ('one2many', 'many2many'):
                    continue
                if total and field.ttype not in ('integer', 'float', 'monetary'):
                    total = False
                result.append((0, 0, {'section': section, 'sequence': seq, 'name': label, 'field_id': field.id,
                                      'total': total}))
            return result

        vals = {
            'name': spec['name'], 'model_id': model.id, 'layout': spec.get('layout', 'record'),
            'output_name': spec.get('output', '{{ object.display_name }}'),
            'multi_mode': 'single' if spec.get('layout') == 'list' else 'sheets',
            'header_line_ids': lines(spec['model'], spec['header'], 'header'),
        }
        if spec.get('title'):
            vals['title'] = spec['title']
        if spec.get('footer'):
            vals['footer'] = spec['footer']
        if spec.get('table'):
            table = Fields._get(spec['model'], spec['table'])
            if table:
                vals['table_field_id'] = table.id
                vals['table_line_ids'] = lines(table.relation, spec['columns'], 'table')
        return vals

    def action_create(self):
        self.ensure_one()
        if self.template not in dict(self._selection()):
            raise UserError((getattr(self.env, "_", None) or _odoo_t)("This report needs an app that is not installed."))
        template = self.env['axs.excel.template'].create(self._template_vals(self.template))
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'axs.excel.template',
            'res_id': template.id,
            'view_mode': 'form',
            'views': [(False, 'form')],
            'target': 'current',
        }
