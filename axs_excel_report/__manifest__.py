{
    'name': 'Excel Report Studio | Excel Templates for Any Document or List',
    'version': '16.0.1.0.0',
    'summary': 'Design Excel reports for any model or fill your own Excel templates with placeholders, '
               'then print them from the form or the list view.',
    'description': """
Excel Report Studio by AXSync
=============================
Excel reports for any Odoo model: a designer that builds the workbook from the fields you choose, or your own
.xlsx templates with placeholders, repeated rows, formulas, pictures, totals and formatting. Print from the Print
menu of a record or of a selection, download, attach or email the file.
""",
    'category': 'Productivity',
    'author': 'AXSync',
    'website': 'https://axsync.com',
    'support': 'hunain@axsync.com',
    'license': 'OPL-1',
    'price': 40.0,
    'currency': 'USD',
    'depends': ['mail'],
    'external_dependencies': {'python': ['openpyxl']},
    'data': [
        'security/security.xml',
        'security/ir.model.access.csv',
        'data/cron.xml',
        'views/excel_template_views.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'axs_excel_report/static/src/scss/excel_report.scss',
        ],
    },
    'images': ['static/description/banner.png'],
    'uninstall_hook': '_uninstall_hook',
    'installable': True,
    'application': True,
}
