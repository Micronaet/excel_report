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

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    #default = fields.Char('Default')
    name = fields.Char('Name', size=64, required=True)
    sequence = fields.Integer('Sequence')
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
    #default = fields.Char('Default')
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
    rgb = fields.char('RGB syntax', size=10, required=True)

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

    font_id = fields.Many2one(
        'excel.report.format.font', 'Font', required=True)
    foreground_id = fields.Many2one('excel.report.format.color', 'Color')
    background_id = fields.Many2one('excel.report.format.color', 'Backgroung')

    height = fields.Integer('Font height', required=True, default=10)
    bold = fields.Boolean('Bold')
    italic = fields.Boolean('Italic')

class ExcelReportFormat(models.Model):
    """ Model name: Inherit for relation: ExcelReportFormat
    """    
    _inherit = 'excel.report.format'
    
    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    style_ids: fields.One2many(
        'excel.report.format.style', 'format_id', 'Style'),
    
class ExcelReport(models.Model):
    """ Model name: Excel Report
    """    
    _name = 'excel.report'
    _description = 'Excel report'
    _order = 'name',

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
    name = fields.Char('Name', size=64, required=True)
    code = fields.Char('Code', size=15, required=True)
    
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
    
    # Excel utility:
    @api.model
    def _create_workbook(self, extension='xlsx'):
        ''' Create workbook in a temp file
        '''
        now = fields.Datetime.now()
        now = now.replace(':', '_').replace('-', '_').replace(' ', '_')
        filename = '/tmp/wb_%s.%s' % (now, extension)
             
        _logger.info('Start create file %s' % filename)
        self._WB = xlsxwriter.Workbook(filename)
        self._WS = {}
        self._filename = filename
        _logger.warning('Created WB and file: %s' % filename)
        
        self.set_format() # setup default format for text used
        self.get_format() # Load database of formats

    @api.model
    def _close_workbook(self, ):
        ''' Close workbook
        '''
        self._WS = {}
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

    @api.model
    def create_worksheet(self, name=False, extension='xlsx'):
        ''' Create database for WS in this module
        '''
        try:
            if not self._WB:
                self._create_workbook(extension=extension)
            _logger.info('Using WB: %s' % self._WB)
        except:
            self._create_workbook(extension=extension)
            
        self._WS[name] = self._WB.add_worksheet(name)
        
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
            'url': '/web/content/excel.writer/%s/b64_file/%s?download=true' % (
                temp_id,
                name_of_file,
                ),
            }

    @api.model
    def merge_cell(self, WS_name, rectangle, default_format=False, data=''):
        ''' Merge cell procedure:
            WS: Worksheet where work
            rectangle: list for 2 corners xy data: [0, 0, 10, 5]
            default_format: setup format for cells
        '''
        rectangle.append(data)        
        if default_format:
            rectangle.append(default_format)            
        self._WS[WS_name].merge_range(*rectangle)
        return 
             
    @api.model
    def write_xls_line(self, WS_name, row, line, default_format=False, col=0):
        ''' Write line in excel file:
            WS: Worksheet where find
            row: position where write
            line: Row passed is a list of element or tuple (element, format)
            default_format: if present replace when format is not present
            
            @return: nothing
        '''
        for record in line:
            if type(record) == bool:
                record = ''
            if type(record) not in (list, tuple):
                if default_format:                    
                    self._WS[WS_name].write(row, col, record, default_format)
                else:    
                    self._WS[WS_name].write(row, col, record)                
            elif len(record) == 2: # Normal text, format
                self._WS[WS_name].write(row, col, *record)
            else: # Rich format TODO
                
                self._WS[WS_name].write_rich_string(row, col, *record)
            col += 1
        return True

    @api.model
    def write_xls_data(self, WS_name, row, col, data, default_format=False):
        ''' Write data in row col position with default_format
            
            @return: nothing
        '''
        if default_format:
            self._WS[WS_name].write(row, col, data, default_format)
        else:    
            self._WS[WS_name].write(row, col, data, default_format)
        return True
        
    @api.model
    def column_width(self, WS_name, columns_w, col=0):
        ''' WS: Worksheet passed
            columns_w: list of dimension for the columns
        '''
        for w in columns_w:
            self._WS[WS_name].set_column(col, col, w)
            col += 1
        return True

    @api.model
    def row_height(self, WS_name, row_list, height=10):
        ''' WS: Worksheet passed
            columns_w: list of dimension for the columns
        '''
        if type(row_list) in (list, tuple):            
            for row in row_list:
                self._WS[WS_name].set_row(row, height)
        else:        
            self._WS[WS_name].set_row(row_list, height)                
        return True
        
    @api.model
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
        ''' Setup 4 element used in normal reporting 
            Every time replace format setup with new database           
        '''
        self._default_format = {
            'title': (title_font, title_size, title_fg),
            'header': (header_font, header_size, header_fg),
            'text': (text_font, text_size, text_fg),
            'number': number_format,
            'border': border,
            }
        _logger.warning('Set format variables: %s' % self._default_format)            
        return
    
    @api.model
    def get_format(self, key=False):  
        ''' Database for format cells
            key: mode of format
            if not passed load database only
        '''
        #try:
        _logger.warning('Set format WB type')
        WB = self._WB # Create with start method
        #except:
            
        F = self._default_format # readability
        
        # Save database in self:
        create = False
        try:
            if not self._wb_format: # raise error if not present
                create = True            
        except:    
            create = True

        if create:    
            self._wb_format = {
                # -------------------------------------------------------------
                # Used when key not present:
                # -------------------------------------------------------------
                'default' : WB.add_format({ # Usually text format
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'align': 'left',
                    }),

                # -------------------------------------------------------------
                #                       TITLE:
                # -------------------------------------------------------------
                'title' : WB.add_format({
                    'bold': True, 
                    'font_name': F['title'][0],
                    'font_size': F['title'][1],
                    'font_color': F['title'][2],
                    'align': 'left',
                    }),
                    
                # -------------------------------------------------------------
                #                       HEADER:
                # -------------------------------------------------------------
                'header': WB.add_format({
                    'bold': True, 
                    'font_name': F['header'][0],
                    'font_size': F['header'][1],
                    'font_color': F['header'][2],
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#cfcfcf', # grey
                    'border': F['border'],
                    #'text_wrap': True,
                    }),

                # -------------------------------------------------------------
                #                       TEXT:
                # -------------------------------------------------------------
                'text': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),                    
                'text_center': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'align': 'center',
                    #'valign': 'vcenter',
                    }),
                'text_right': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                    
                'text_total': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#DDDDDD',
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True,
                    }),

                # --------------
                # Text BG color:
                # --------------
                # No bold:
                'bg_normal_white': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#FFFFFF',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_normal_red': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#ffc6af',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_normal_green': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#b1f9c1',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),

                'bg_normal_white_number': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#FFFFFF',
                    'align': 'right',
                    'num_format': F['number'],
                    }),                
                'bg_normal_red_number': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#ffc6af',
                    'align': 'right',
                    'num_format': F['number'],
                    }),
                'bg_normal_green_number': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#b1f9c1',
                    'align': 'right',
                    'num_format': F['number'],
                    }),


                # Bold
                'bg_white': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#FFFFFF',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_blue': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#c4daff',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_red': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#ffc6af',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_green': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#b1f9c1',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),
                'bg_yellow': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#fffec1',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),                
                'bg_orange': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#fcdebd',
                    'align': 'left',
                    #'valign': 'vcenter',
                    }),                
                'bg_red_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#ffc6af',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),
                'bg_green_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'bg_color': '#b1f9c1',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),
                'bg_yellow_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#fffec1',##ffff99',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),                
                'bg_orange_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#fcdebd',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),                
                'bg_white_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#FFFFFF',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),                
                'bg_blue_number': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'font_color': 'black',
                    'bg_color': '#c4daff',##ffff99',
                    'align': 'right',
                    'num_format': F['number'],
                    #'valign': 'vcenter',
                    }),                

                # TODO remove?
                'bg_order': WB.add_format({
                    'bold': True, 
                    'bg_color': '#cc9900',
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'num_format': F['number'],
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),

                # --------------
                # Text FG color:
                # --------------
                'text_black': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': 'black',
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True
                    }),
                'text_blue': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': 'blue',
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True
                    }),
                'text_red': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': '#ff420e',
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True
                    }),
                'text_green': WB.add_format({
                    'font_color': '#328238', ##99cc66
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True
                    }),
                'text_grey': WB.add_format({
                    'font_color': '#eeeeee',
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True
                    }),                
                'text_wrap': WB.add_format({
                    'font_color': 'black',
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'align': 'left',
                    'valign': 'vcenter',
                    #'text_wrap': True,
                    }),

                # -------------------------------------------------------------
                #                       NUMBER:
                # -------------------------------------------------------------
                'number': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),

                # ----------------
                # Number FG color:
                # ----------------
                'number_black': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'font_color': 'black',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                'number_blue': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'font_color': 'blue',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                'number_grey': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'font_color': 'grey',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                'number_red': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'font_color': 'red',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                'number_green': WB.add_format({
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'border': F['border'],
                    'num_format': F['number'],
                    'font_color': 'green',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),

                'number_total': WB.add_format({
                    'bold': True, 
                    'font_name': F['text'][0],
                    'font_size': F['text'][1],
                    'font_color': F['text'][2],
                    'border': F['border'],
                    'num_format': F['number'],
                    'bg_color': '#DDDDDD',
                    'align': 'right',
                    #'valign': 'vcenter',
                    }),
                }
        
        # Return format or default one's
        if key:
            return self._wb_format.get(
                key, 
                self._wb_format.get('default'),
                )
        else:
            return True    
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
