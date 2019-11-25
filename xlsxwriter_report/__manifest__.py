# Copyright 2019  Micronaet SRL (<http://www.micronaet.it>).
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl.html).

{
    'name': 'XLSX report',
    'version': '11.0.2.0.0',
    'category': 'Report',
    'description': '''
        Template for python xlsxwriter report
        ''',
    'summary': 'Excel, utility, report',
    'author': 'Micronaet S.r.l. - Nicola Riolini',
    'website': 'http://www.micronaet.it',
    'license': 'AGPL-3',
    'depends': [
        'report_xlsx',
    ],
    'data': [
        'security/ir.model.access.csv',
        'views/excel_report_view.xml',
        'data/color_data.xml',
        'data/border_data.xml',
        'data/font_data.xml',
        'data/page_data.xml',
        'data/format_data.xml',
        'data/style_data.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'application': False,
    'installable': True,
    'auto_install': False,
}
