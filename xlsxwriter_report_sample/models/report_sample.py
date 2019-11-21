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

class ResPartnerExcelReportWizard(models.TransientModel):
    """ Model name: StockPicking
    """
    _name = 'res.partner.excel.report.wizard'
    _description = 'Extract partner report'
    
    # -------------------------------------------------------------------------
    #                            COLUMNS:
    # -------------------------------------------------------------------------    
    country_id = fields.Many2one('res.country', 'Country')
    # -------------------------------------------------------------------------    

    @api.multi
    def excel_partner_report(self, ):
        ''' Extract Excel PFU report
        '''
        report_pool = self.env['excel.report']
        partner_pool = self.env['res.partner']
        
        # ---------------------------------------------------------------------
        # Wizard parameters:
        # ---------------------------------------------------------------------
        country = self.country_id
                
        # ---------------------------------------------------------------------
        # Collect data:
        # ---------------------------------------------------------------------
        domain = []
        if country:
            domain.append(('country_id', '=', country.id))
        partners = partner_pool.search(domain)

        # ---------------------------------------------------------------------
        #                          EXTRACT EXCEL:
        # ---------------------------------------------------------------------
        # Excel file configuration:
        title = ('Partner list', )
        header = ('Name', 'City', 'Country')            
        column_width = (40, 30, 20)    

        # ---------------------------------------------------------------------
        # Write detail:
        # ---------------------------------------------------------------------        
        ws_name = 'Partner' # Worksheet name:
        report_pool.create_worksheet(ws_name, format_code='DEFAULT')
        report_pool.column_width(ws_name, column_width)

        # Title:
        row = 0
        report_pool.write_xls_line(ws_name, row, title, style_code='title')
        
        # Header:
        row += 1
        report_pool.write_xls_line(ws_name, row, header, style_code='header')
        
        for partner in sorted(partners, key=lambda x: x.name):
            # Data line:
            row += 1
            
            # Setup color line:
            if partner.country_id.name:
               style_code = 'text'
            else:
               style_code = 'text_error'
               
            report_pool.write_xls_line(ws_name, row, (
                partner.name,
                partner.city or '',
                partner.country_id.name or '',
                ), style_code=style_code)
                
        # ---------------------------------------------------------------------
        # Save file:
        # ---------------------------------------------------------------------
        return report_pool.return_attachment('Report_Partner')

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
