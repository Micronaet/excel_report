# Copyright 2019  Micronaet SRL (<http://www.micronaet.it>).
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl.html).
import os
import sys
import logging
import odoo
from odoo import api, fields, models, tools, exceptions, SUPERUSER_ID
from odoo.addons import decimal_precision as dp
from odoo.tools.translate import _


_logger = logging.getLogger(__name__)

class ProductProductExcelReportWizard(models.TransientModel):
    _name = 'product.product.excel.report.wizard'
    _description = 'Extract product report'

    category_id = fields.Many2one('product.category', 'Category')

    @api.multi
    def excel_partner_report(self, ):
        report_pool = self.env['excel.report']
        product_pool = self.env['product.product']
        category = self.category_id
        # Collect data:
        domain = []
        if category:
            domain.append(('categ_id', '=', category.id))
        products = product_pool.search(domain)

        # Excel file configuration:
        title = ('Product list (red line = product no price)', )
        header = ('Name', 'Code', 'Category', 'Tax', 'Weight', 'List price', )            
        column_width = (40, 30, 20, 15, 10, 10)
        total_columns = (4, 5)  # Columns used for total

        ws_name = _('Product')  # Worksheet name
        report_pool.create_worksheet(ws_name, format_code='DEFAULT')
        report_pool.column_width(ws_name, column_width)

        # Title:
        row = 0
        report_pool.write_xls_line(ws_name, row, title, style_code='title')
        
        # Merge title cell (first row, N cols):
        report_pool.merge_cell(ws_name, [row, 0, row, len(header) -1])

        # Header:
        row += 1
        report_pool.write_xls_line(ws_name, row, header, style_code='header')
        
        # Set auto-filter (where needed: category, tax)
        report_pool.autofilter(ws_name, [row, 2, row, 3])

        # Data lines:
        for product in sorted(products, key=lambda x: x.name):
            row += 1
            # Setup color line (red = product empty price):
            if product.list_price:
                color_style_text = 'text'
                color_style_number = 'number'
            else:
                color_style_text = 'text_error'
                color_style_number = 'number_error'

            # Write data:
            report_pool.write_xls_line(ws_name, row, (
                product.name,
                product.default_code or '',
                product.categ_id.name or '',
                product.taxes_id.name or '',
                (product.weight, color_style_number),
                (product.list_price, color_style_number),
                ), style_code=color_style_text, total_columns=total_columns)

        # Write total line:
        row += 1
        report_pool.write_total_xls_line(ws_name, row, total_columns, style_code='number_total')

        # Save file:
        return report_pool.return_attachment('Report_Product')
