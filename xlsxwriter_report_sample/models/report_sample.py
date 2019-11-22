#!/usr/bin/python
# -*- coding: utf-8 -*-
###############################################################################
#
# ODOO (ex OpenERP) 
# Open Source Management Solution
# Copyright (C) 2001-2015 Micronaet S.r.l. (<https://micronaet.com>)
# Developer: Nicola Riolini @thebrush (<https://it.linkedin.com/in/thebrush>)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
# See the GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################

import os
import sys
import logging
import odoo
from odoo import api, fields, models, tools, exceptions, SUPERUSER_ID
from odoo.addons import decimal_precision as dp
from odoo.tools.translate import _


_logger = logging.getLogger(__name__)

class ProductProductExcelReportWizard(models.TransientModel):
    """ Model name: StockPicking
    """
    _name = 'product.product.excel.report.wizard'
    _description = 'Extract product report'
    
    # -------------------------------------------------------------------------
    #                            COLUMNS:
    # -------------------------------------------------------------------------    
    category_id = fields.Many2one('product.category', 'Category')
    # -------------------------------------------------------------------------    

    @api.multi
    def excel_partner_report(self, ):
        ''' Extract Excel PFU report
        '''
        report_pool = self.env['excel.report']
        product_pool = self.env['product.product']
        
        # ---------------------------------------------------------------------
        # Wizard parameters:
        # ---------------------------------------------------------------------
        category = self.category_id
                
        # ---------------------------------------------------------------------
        # Collect data:
        # ---------------------------------------------------------------------
        domain = []
        if category:
            domain.append(('categ_id', '=', category.id))
        products = product_pool.search(domain)

        # ---------------------------------------------------------------------
        #                          EXTRACT EXCEL:
        # ---------------------------------------------------------------------
        # Excel file configuration:
        title = ('Product list (red line = product no price)', )
        header = ('Name', 'Code', 'Category', 'Tax', 'Weight', 'List price', )            
        column_width = (40, 30, 20, 15, 10, 10)    

        # ---------------------------------------------------------------------
        # WRITE DATA:
        # ---------------------------------------------------------------------        
        ws_name = _('Product') # Worksheet name
        report_pool.create_worksheet(ws_name, format_code='DEFAULT')
        report_pool.column_width(ws_name, column_width)

        # ---------------------------------------------------------------------        
        # Title:
        row = 0
        report_pool.write_xls_line(ws_name, row, title, style_code='title')
        
        # Merge title cell (first row, N cols):
        report_pool.merge_cell(ws_name, [row, 0, row, len(header) -1])
        
        # ---------------------------------------------------------------------        
        # Header:
        row += 1
        report_pool.write_xls_line(ws_name, row, header, style_code='header')
        
        # Set autofilter (where needed: category, tax)
        report_pool.autofilter(ws_name, [row, 2, row, 3])

        # ---------------------------------------------------------------------        
        # Data lines:
        #total_columns = (3, 4)
        for product in sorted(products, key=lambda x: x.name):
            row += 1
             
            # -----------------------------------------------------------------
            # Setup color line (red = product empty price):
            # -----------------------------------------------------------------
            if product.list_price:
               color_style_text = 'text'
               color_style_number = 'number'
            else:
               color_style_text = 'text_error'
               color_style_number = 'number_error'
               
            # -----------------------------------------------------------------
            # Write data:
            # -----------------------------------------------------------------
            report_pool.write_xls_line(ws_name, row, (
                product.name,
                product.default_code or '',
                product.categ_id.name or '',
                product.taxes_id.name or '',
                (product.weight, color_style_number),
                (product.list_price, color_style_number),
                ), style_code=color_style_text) #total_columns=total_columns
                
        # ---------------------------------------------------------------------
        # Save file:
        # ---------------------------------------------------------------------
        return report_pool.return_attachment('Report_Product')

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
