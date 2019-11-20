# -*- coding: utf-8 -*-
###############################################################################
#
#    Copyright (C) 2001-2014 Micronaet SRL (<http://www.micronaet.it>).
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as published
#    by the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################
import os
import sys
import logging
import base64
import xlsxwriter
import shutil
import openerp
import logging

from openerp import models, fields, api
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openerp.tools.translate import _
from openerp.tools import (
    DEFAULT_SERVER_DATE_FORMAT, 
    DEFAULT_SERVER_DATETIME_FORMAT, 
    DATETIME_FORMATS_MAP, 
    float_compare,
    )


_logger = logging.getLogger(__name__)

class ExcelReportFormatPage(models.Model):
    """ Model name: ExcelReportFormatPage
    """    
    _name = 'excel.report.format.page'
    _description = 'Excel report'
    _order = 'sequence,name'

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    #default = fields.Char('Default')
    index = fields.Integer('Index', required=True)
    name = fields.Char('Name', size=64, required=True)
    sequence = fields.Integer('Sequence')
    paper_size = fields.Char('Paper size', size=40)
    # dimension
    # note

class ExcelReportFormat(models.Model):
    """ Model name: ExcelReportFormat
    """    
    _name = 'excel.report.format'
    _description = 'Excel report'
    
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    #default = fields.Boolean('Default')
    name = fields.Char('Name', size=64, required=True)
    code = fields.Char('Code', size=15, required=True)
    page_id = fields.Many2one(
        'excel.report.format.page', 'Page', required=True)
    
    margin_top = fields.Integer('Margin Top')
    margin_bottom = fields.Integer('Margin Bottom')
    margin_left = fields.Integer('Margin Left')
    margin_right = fields.Integer('Margin Right')

    # TODO header, footer

class ExcelReportFormatFont(models.Model):
    """ Model name: ExcelReportFormatFont
    """    
    _name = 'excel.report.format.font'
    _description = 'Excel format font'
        
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    name = fields.Char('Font name', size=64, required=True)

class ExcelReportFormatColor(models.Model):
    """ Model name: ExcelReportFormatColor
    """    
    _name = 'excel.report.format.color'
    _description = 'Excel format color'
        
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    name = fields.Char('Color name', size=64, required=True)
    rgb = fields.Char('RGB syntax', size=10, required=True)

# TODO class With numer format

class ExcelReportFormatStyle(models.Model):
    """ Model name: ExcelReportFormat
    """    
    _name = 'excel.report.format.style'
    _description = 'Excel format style'
    
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    name = fields.Char('Name', size=64, required=True)
    code = fields.Char('Code', size=15, required=True)
    format_id = fields.Many2one('excel.report.format', 'Format')

    # -------------------------------------------------------------------------
    # Font:
    font_id = fields.Many2one(
        'excel.report.format.font', 'Font', required=True, 
        help='Remember to use standard fonts, need to be installed on PC!')
    foreground_id = fields.Many2one('excel.report.format.color', 'Color')
    background_id = fields.Many2one('excel.report.format.color', 'Backgroung')

    height = fields.Integer('Font height', required=True, default=10)
    
    # -------------------------------------------------------------------------
    # Type:
    bold = fields.Boolean('Bold')
    italic = fields.Boolean('Italic')
    
    # -------------------------------------------------------------------------
    # Border:
    border_top = fields.Integer('Border top')
    border_bottom = fields.Integer('Border bottom')
    border_left = fields.Integer('Border left')
    border_right = fields.Integer('Border right')

    # -------------------------------------------------------------------------
    # Border color
    border_color_top_id = fields.Many2one(
        'excel.report.format.color', 'Border top color')
    border_color_bottom_id = fields.Many2one(
        'excel.report.format.color', 'Border bottom color')
    border_color_left_id = fields.Many2one(
        'excel.report.format.color', 'Border left color')
    border_color_right_id = fields.Many2one(
        'excel.report.format.color', 'Border right color')
    
    # TODO: 
    # align
    # valign
    # wrap
    # format

class ExcelReportFormat(models.Model):
    """ Model name: Inherit for relation: ExcelReportFormat
    """    
    _inherit = 'excel.report.format'
    
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    style_ids = fields.One2many(
        'excel.report.format.style', 'format_id', 'Style')
    
class ExcelReport(models.TransientModel):
    """ Model name: Excel Report
    """    
    _name = 'excel.report'
    _description = 'Excel report'
    _order = 'name'

    # -------------------------------------------------------------------------
    # Computed fields:
    # -------------------------------------------------------------------------
    @api.one
    def _get_template(self):
        ''' Computed fields: B64 file from file content
        '''
        try:
            origin = self.fullname
            self.b64_file = base64.b64encode(open(origin, 'rb').read())
        except:
            self.b64_file = False    

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    #name = fields.Char('Name', size=64, required=True)
    #code = fields.Char('Code', size=15, required=True)
    
    b64_file = fields.Binary('B64 file', compute='_get_template')
    fullname = fields.Text('Fullname of file')
    
    # -------------------------------------------------------------------------
    #                                   UTILITY:
    # -------------------------------------------------------------------------
    @api.model
    def clean_filename(self, destination):
        ''' Clean char that generate error
        '''
        destination = destination.replace('/', '_').replace(':', '_')
        if not(destination.endswith('xlsx') or destination.endswith('xls')):
            destination = '%s.xlsx' % destination
        return destination    
        
    # Format utility:
    @api.model
    def format_date(self, value):
        ''' Format hour DD:MM:YYYY
        '''
        if not value:
            return ''
        return '%s/%s/%s' % (
            value[8:10],
            value[5:7],
            value[:4],
            )

    @api.model
    def format_hour(self, value, hhmm_format=True, approx=0.001, 
            zero_value='0:00'):
        ''' Format hour HH:MM
        '''
        if not hhmm_format:
            return value
            
        if not value:
            return zero_value
            
        value += approx    
        hour = int(value)
        minute = int((value - hour) * 60)
        return '%d:%02d' % (hour, minute) 
    
    # -------------------------------------------------------------------------
    #                              Excel utility:
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    # Workbook:
    # -------------------------------------------------------------------------
    @api.model
    def _create_workbook(self, extension='xlsx'):
        ''' Create workbook in a temp file
        '''
        now = fields.Datetime.now()
        now = now.replace(':', '_').replace('-', '_').replace(' ', '_')
        filename = '/tmp/wb_%s.%s' % (now, extension) # TODO better!

        _logger.info('Start create file %s' % filename)
        self._WB = xlsxwriter.Workbook(filename)
        self._WS = {}
        self._style = {} # Style for every WS

        self._filename = filename
        _logger.warning('Created WB on file: %s' % filename)

    @api.model
    def _close_workbook(self, ):
        ''' Close workbook
        '''
        self._WS = {}
        self._style = {}
        self._wb_format = False
        
        try:
            self._WB.close()            
        except:            
            _logger.error('Error closing WB')    
        self._WB = False # remove object in instance

    @api.model
    def close_workbook(self, ):
        ''' Close workbook
        '''
        return self._close_workbook()

    # -------------------------------------------------------------------------
    # Worksheet:
    # -------------------------------------------------------------------------
    @api.model
    def create_worksheet(self, name=False, format_code='', extension='xlsx'):
        ''' Create database for WS in this module
        '''
        try:
            if not self._WB:
                self._create_workbook(extension=extension)
            _logger.info('Using WB: %s' % self._WB)
        except:
            self._create_workbook(extension=extension)
            
        self._WS[name] = self._WB.add_worksheet(name)
        
        # Setup Format:
        if format_code:
            self._load_format_code(name, format_code)
            
    # -------------------------------------------------------------------------
    # Format:
    # -------------------------------------------------------------------------
    @api.model
    def _load_format_code(self, name, format_code):
        ''' Setup format parameters and syles
        '''
        format_pool = self.env['excel.report.format']
        formats = format_pool.search([('code', '=', format_code)])
        ws = self._WS[name]
        if formats:
            current_format = formats[0]
            _logger.info('Format selected: %s' % format_code)
            
            # Setup page:
            page_id = current_format.page_id
            if page_id:
                # -------------------------------------------------------------
                # Set page:
                # -------------------------------------------------------------
                ws.set_paper(page_id.index)

                # -------------------------------------------------------------
                # Set orientation: 
                # -------------------------------------------------------------
                # set_landscape set_portrait
                if True:     # TODO paramter to add!
                    ws.set_landscape() 
                else:    
                    ws.set_portrait()
            
                # -------------------------------------------------------------
                # Setup Margin
                # -------------------------------------------------------------
                ws.set_margins(
                    left=current_format.margin_left, 
                    right=current_format.margin_right, 
                    top=current_format.margin_top, 
                    bottom=current_format.margin_bottom,
                    )
                    
                # -------------------------------------------------------------
                # Load Styles:
                # -------------------------------------------------------------
                if name not in self._style:
                    # Every page use own style (can use different format)
                    self._style[name] = {}

                for style in current_format.style_ids:
                    # Create new style and add
                    self._style[name][style.code] = self._WB.add_format({
                        'bold': style.bold, 
                        'font_name': style.font_id.name,
                        'font_size': style.height,
                        'font_color': style.foreground_id.rgb,
                        'bg_color': style.background_id.rgb,
                        #'align': 'center',
                        #'valign': 'vcenter',
                        #'border': F['border'],
                        # border color
                        # italic
                        #'text_wrap': True,
                        #'num_format': F['number'],
                        }),
            
        else:    
            _logger.info('Format not found: %s, use nothing: %s' % format_code)
        
        
    # -------------------------------------------------------------------------
    # Sheet setup:
    # -------------------------------------------------------------------------
    @api.model
    def column_width(self, ws_name, columns_w, col=0):
        ''' WS: Worksheet passed
            columns_w: list of dimension for the columns
        '''
        for w in columns_w:
            self._WS[ws_name].set_column(col, col, w)
            col += 1
        return True

    @api.model
    def row_height(self, ws_name, row_list, height=10):
        ''' WS: Worksheet passed
            columns_w: list of dimension for the columns
        '''
        if type(row_list) in (list, tuple):            
            for row in row_list:
                self._WS[ws_name].set_row(row, height)
        else:        
            self._WS[ws_name].set_row(row_list, height)                
        return True

    @api.model
    def merge_cell(self, ws_name, rectangle, style=False, data=''):
        ''' Merge cell procedure:
            WS: Worksheet where work
            rectangle: list for 2 corners xy data: [0, 0, 10, 5]
            style: setup format for cells
        '''
        rectangle.append(data)        
        if style:
            rectangle.append(style)            
        self._WS[ws_name].merge_range(*rectangle)
        return 

    # -------------------------------------------------------------------------
    # Miscellaneous operations (called directly):
    # -------------------------------------------------------------------------
    @api.model
    def write_xls_line(self, ws_name, row, line, style_code=False, col=0):
        ''' Write line in excel file:
            WS: Worksheet where find
            row: position where write
            line: Row passed is a list of element or tuple (element, format)
            style: if present replace when format is not present
            
            @return: nothing
        '''
        style = self._WS(ws_name).get(style_code)
        for record in line:
            if type(record) == bool:
                record = ''
            if type(record) not in (list, tuple):
                if style: # needed?               
                    self._WS[ws_name].write(row, col, record, style)
                else:    
                    self._WS[ws_name].write(row, col, record)                
            elif len(record) == 2: # Normal text, format
                self._WS[ws_name].write(row, col, *record)
            else: # Rich format TODO                
                self._WS[ws_name].write_rich_string(row, col, *record)
            col += 1
        return True

    # TODO USED?
    """
    @api.model
    def write_xls_data(self, ws_name, row, col, data, style=False):
        ''' Write data in row col position with style (used?)
            
            @return: nothing
        '''
        if style:
            self._WS[ws_name].write(row, col, data, style)
        else:    
            self._WS[ws_name].write(row, col, data, style)
        return True"""
        
    # -------------------------------------------------------------------------
    # Return operation:
    # -------------------------------------------------------------------------
    @api.model
    def send_mail_to_group(self,
            group_name,
            subject, body, filename, # Mail data
            ):
        ''' Send mail of current workbook to all partner present in group 
            passed
            group_name: use format module_name.group_id
            subject: mail subject
            body: mail body
            filename: name of xlsx attached file
        '''
        # Send mail with attachment:
        
        # Pool used
        group_pool = self.env['res.groups']
        model_pool = self.env['ir.model.data']
        thread_pool = self.env['mail.thread']

        self._close_workbook() # Close before read file
        attachments = [(
            filename, 
            open(self._filename, 'rb').read(), # Raw data
            )]

        group = group_name.split('.')
        groups_id = model_pool.get_object_reference(
            cr, uid, group[0], group[1])[1]    
        partner_ids = []
        for user in group_pool.browse(group_id).users:
            partner_ids.append(user.partner_id.id)
            
        thread_pool = self.env['mail.thread']
        thread_pool.message_post(False, 
            type='email', 
            body=body, 
            subject=subject,
            partner_ids=[(6, 0, partner_ids)],
            attachments=attachments, 
            )
        self._close_workbook() # if not closed maually        

    @api.model
    def save_file_as(self, destination):
        ''' Close workbook and save in another place (passed)
        '''
        _logger.warning('Save file as: %s' % destination)        
        origin = self._filename
        self._close_workbook() # if not closed maually
        shutil.copy(origin, destination)
        return True

    @api.model
    def save_binary_xlsx(self, binary):
        ''' Save binary data passed as file temp (returned)
        '''
        b64_file = base64.decodestring(binary)
        fields.Datetime.now()
        filename = \
            '/tmp/file_%s.xlsx' % now.replace(':', '_').replace('-', '_')
        f = open(filename, 'wb')
        f.write(b64_file)
        f.close()
        return filename

    @api.model
    def return_attachment(self, name, name_of_file=False):
        ''' Return attachment passed
            name: Name for the attachment
            name_of_file: file name downloaded
            php: paremeter if activate save_as module for 7.0 (passed base srv)
            context: context passed
        '''
        if not name_of_file:
            now = fields.Datetime.now()
            now = now.replace('-', '_').replace(':', '_') 
            #name_of_file = '/tmp/report_%s.xlsx' % now
            name_of_file = 'report_%s.xlsx' % now
        self._close_workbook() # if not closed maually
        _logger.info('Return XLSX file: %s' % self._filename)
        
        # TODO is necessary?
        temp_id = self.create({
            'fullname': self._filename,
            }).id
        
        return {
            'type' : 'ir.actions.act_url',
            'name': name,
            'url': '/web/content/excel.report/%s/b64_file/%s?download=true' % (
                temp_id,
                name_of_file,
                ),
            }

        
    '''@api.model
    def set_format(    
            self, 
            # Title:
            title_font='Courier 10 pitch', title_size=11, title_fg='black', 
            # Header:
            header_font='Courier 10 pitch', header_size=9, header_fg='black',
            # Text:
            text_font='Courier 10 pitch', text_size=9, text_fg='black',
            # Number:
            number_format='#,##0.#0',
            # Layout:
            border=1,
            ):
        '''
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
