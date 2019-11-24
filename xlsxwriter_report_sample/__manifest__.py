# Copyright 2019  Micronaet SRL (<http://www.micronaet.it>).
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl.html).

{
    'name': 'Excel report Sample',
    'version': '1.0',
    'category': 'Reporting',
    'sequence': 5,
    'summary': 'Report, Example, Startup',
    'website': 'https://micronaet.com',
    'depends': [
        'base',
        'product',
        'xlsxwriter_report',
    ],
    'data': [
        'views/report_sample_view.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
