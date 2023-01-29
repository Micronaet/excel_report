# Copyright 2020  Micronaet SRL (<http://www.micronaet.it>).
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl.html).

import io
import xlsxwriter
import logging
import base64
import shutil
from odoo import models, fields, api
from xlsxwriter.utility import xl_rowcol_to_cell


_logger = logging.getLogger(__name__)


class ExcelReportFormatPage(models.Model):
    """ Model name: ExcelReportFormatPage
    """
    _name = 'excel.report.format.page'
    _description = 'Excel report'
    _order = 'index'

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    # default = fields.Char('Default')
    index = fields.Integer('Index', required=True)
    name = fields.Char('Name', size=64, required=True)
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
    # default = fields.Boolean('Default')
    name = fields.Char('Name', size=64, required=True)
    code = fields.Char('Code', size=15, required=True)
    page_id = fields.Many2one(
        'excel.report.format.page', 'Page', required=True)
    row_height = fields.Integer(
        'Row height',
        help='Usually setup in style, if not take this default value!')

    margin_top = fields.Float('Margin Top', digits=(16, 3), default=0.25)
    margin_bottom = fields.Float('Margin Bottom', digits=(16, 3), default=0.25)
    margin_left = fields.Float('Margin Left', digits=(16, 3), default=0.25)
    margin_right = fields.Float('Margin Right', digits=(16, 3), default=0.25)

    orientation = fields.Selection([
        ('portrait', 'Portrait'),
        ('landscape', 'Landscape'),
        ], 'Orientation', default='portrait')

    # todo header, footer


class ExcelReportFormatFont(models.Model):
    """ Model name: ExcelReportFormatFont
    """
    _name = 'excel.report.format.font'
    _description = 'Excel format font'

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    name = fields.Char('Font name', size=64, required=True)


class ExcelReportFormatBorder(models.Model):
    """ Model name: ExcelReportFormatColor
    """
    _name = 'excel.report.format.border'
    _description = 'Excel format border'

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    name = fields.Char('Color name', size=64, required=True)
    index = fields.Integer('Index', required=True)
    weight = fields.Integer('Weight')
    style = fields.Char('Style', size=20)


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

# todo class With number format


class ExcelReportFormatStyle(models.Model):
    """ Model name: ExcelReportFormat
    """
    _name = 'excel.report.format.style'
    _description = 'Excel format style'

    name = fields.Char('Name', size=64, required=True)
    code = fields.Char('Code', size=15, required=True)
    format_id = fields.Many2one('excel.report.format', 'Format')
    row_height = fields.Integer(
        'Row height',
        help='If present use this, instead format value!')
    font_id = fields.Many2one(
        'excel.report.format.font', 'Font', required=True,
        help='Remember to use standard fonts, need to be installed on PC!')
    foreground_id = fields.Many2one('excel.report.format.color', 'Color')
    background_id = fields.Many2one('excel.report.format.color', 'Background')

    height = fields.Integer('Font height', required=True, default=10)

    # Type:
    bold = fields.Boolean('Bold')
    italic = fields.Boolean('Italic')
    num_format = fields.Char('Number format', size=20)  # , default='#,##0.00')

    # -------------------------------------------------------------------------
    # Border:
    border_top_id = fields.Many2one(
        'excel.report.format.border', 'Border top')
    border_bottom_id = fields.Many2one(
        'excel.report.format.border', 'Border bottom')
    border_left_id = fields.Many2one(
        'excel.report.format.border', 'Border left')
    border_right_id = fields.Many2one(
        'excel.report.format.border', 'Border right')

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

    # -------------------------------------------------------------------------
    # Alignment:
    align = fields.Selection([
        ('left', 'Left'),
        ('center', 'Center'),
        ('right', 'Right'),
        ('fill', 'Fill'),
        ('justify', 'Justify'),
        ('center_across', 'Center across'),
        ('distributed', 'Distributed'),
        ], 'Horizontal alignment', default='left')

    valign = fields.Selection([
        ('top', 'Top'),
        ('vcenter', 'Middle'),
        ('bottom', 'Bottom'),
        ('vjustify', 'Justify'),
        ('vdistributed', 'Distribuited'),
        ], 'Vertical alignment', default='vcenter')
    # todo:
    # wrap
    # format


class ExcelReportFormatInherit(models.Model):
    """ Model name: Inherit for relation: ExcelReportFormat
    """
    _inherit = 'excel.report.format'

    # -------------------------------------------------------------------------
    #                                   COLUMNS:
    # -------------------------------------------------------------------------
    style_ids = fields.One2many(
        'excel.report.format.style', 'format_id', 'Style')


class ExcelReport(models.TransientModel):
    """ Excel Report Wizard
    """
    _name = 'excel.report'
    _description = 'Excel report'
    _order = 'name'

    def create(self, vals):
        """ Generate new WB when create record
        """
        if 'fullname' in vals:
            fullname = vals['fullname']
            vals['fullname'] = fullname
        else:
            now = fields.Datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
            fullname = '/tmp/wb_%s.xlsx' % now  # todo better!
        record = super().create(vals)

        # Update record context with Excel Workbook parameters:
        workbook = {
            'WB': xlsxwriter.Workbook(fullname),
            'WS': {},
            'style': {},  # Style for every WS
            'total': {},  # Array for total line (one for ws)
            'row_height': {},
        }
        return record.with_context(workbook=workbook).browse(record.id)

    def get_b64_from_filename(self, workbook):
        """ Read filename for workbook and return binary
        """
        self.ensure_one()
        try:
            origin = workbook['fullname']
            self.b64_file = base64.b64encode(open(origin, 'rb').read())
        except:
            self.b64_file = False

    b64_file = fields.Binary('B64 file')
    fullname = fields.Text('Fullname of file')

    def clean_filename(self, destination):
        destination = destination.replace('/', '_').replace(':', '_')
        if not(destination.endswith('xlsx') or destination.endswith('xls')):
            destination = '%s.xlsx' % destination
        return destination

    # Format utility:
    def format_date(self, value):
        # Format hour DD:MM:YYYY
        if not value:
            return ''
        return '%s/%s/%s' % (
            value[8:10],
            value[5:7],
            value[:4],
            )

    def format_hour(
            self, value, hhmm_format=True, approx=0.001,
            zero_value='0:00'):
        # Format hour HH:MM
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
    def create_workbook(self, extension='xlsx'):
        """ Create workbook in a temp file
        """
        now = fields.Datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
        filename = '/tmp/wb_%s.%s' % (now, extension)  # todo better!

        _logger.info('Start create file %s' % filename)
        workbook = {
            'WB': xlsxwriter.Workbook(filename),
            'WS': {},
            'style': {},  # Style for every WS
            'total': {},  # Array for total line (one for ws)
            'row_height': {},
            'filename': filename,
        }
        _logger.warning('Created WB on file: %s' % filename)
        return workbook

    def close_workbook(self, workbook):
        """ Close workbook
        """
        try:
            workbook['WB'].close()
        except:
            _logger.error('Error closing WB')
        del workbook

    # -------------------------------------------------------------------------
    # Worksheet:
    # -------------------------------------------------------------------------
    def create_worksheet(self, workbook, name=False, format_code=''):
        """ Create database for WS in this module
        """
        workbook['WS'][name] = workbook['WB'].add_worksheet(name)
        workbook['style'][name] = {}
        workbook['total'][name] = False  # Reset total
        # todo subtotal

        # ---------------------------------------------------------------------
        # Setup Format (every new sheet):
        # ---------------------------------------------------------------------
        if format_code:
            self.load_format_code(workbook, name, format_code)

    # -------------------------------------------------------------------------
    # Format:
    # -------------------------------------------------------------------------
    def load_format_code(self, workbook, name, format_code):
        """ Setup format parameters and styles
        """
        format_pool = self.env['excel.report.format']
        formats = format_pool.search([('code', '=', format_code)])
        ws = workbook['WS'][name]
        if formats:
            current_format = formats[0]
            _logger.info('Format selected: %s' % format_code)

            # Setup page:
            row_height = current_format.row_height or False  # default over.
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
                if current_format.orientation == 'landscape':
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
                if name not in workbook['style']:
                    # Every page use own style (can use different format)
                    workbook['style'][name] = {}

                for style in current_format.style_ids:
                    # Create new style and add
                    workbook['style'][name][style.code] = \
                        workbook['WB'].add_format({
                            'font_name': style.font_id.name,
                            'font_size': style.height,
                            'font_color': style.foreground_id.rgb,

                            'bold': style.bold,
                            'italic': style.italic,

                            # -------------------------------------------------
                            # Border:
                            # -------------------------------------------------
                            # Mode:
                            'bottom': style.border_bottom_id.index or 0,
                            'top': style.border_top_id.index or 0,
                            'left': style.border_left_id.index or 0,
                            'right': style.border_right_id.index or 0,

                            # Color:
                            'bottom_color':
                                style.border_color_bottom_id.rgb or '',
                            'top_color':
                                style.border_color_top_id.rgb or '',
                            'left_color':
                                style.border_color_left_id.rgb or '',
                            'right_color':
                                style.border_color_right_id.rgb or '',

                            'bg_color': style.background_id.rgb,

                            'align': style.align,
                            'valign': style.valign,
                            'num_format': style.num_format or '',
                            # 'text_wrap': True,
                            # locked
                            # hidden
                            })

                    # Save row height for this style:
                    workbook['row_height'][workbook['style'][name][
                        style.code]] = style.row_height or row_height
        else:
            _logger.info('Format not found: %s, use nothing: %s' % format_code)

    # -------------------------------------------------------------------------
    # Sheet setup:
    # -------------------------------------------------------------------------
    def column_width(self, workbook, ws_name, columns_w, col=0):
        """ WS: Worksheet passed
            columns_w: list of dimension for the columns
        """
        for w in columns_w:
            workbook['WS'][ws_name].set_column(col, col, w)
            col += 1
        return True

    def column_hidden(self, workbook, ws_name, columns_w):
        """ WS: Worksheet passed
            columns_w: list of dimension for the columns
        """
        for col in columns_w:
            workbook['WS'][ws_name].set_column(
                col, col, None, None, {'hidden': True})
        return True

    def row_hidden(self, workbook, ws_name, rows_w):
        """ WS: Worksheet passed
            columns_w: list of dimension for the columns
        """
        for row in rows_w:
            workbook['WS'][ws_name].set_row(
                row, None, None, {'hidden': True})
        return True

    def row_height(self, workbook, ws_name, row_list, height=15):
        """ WS: Worksheet passed
            columns_w: list of dimension for the columns
        """
        if type(row_list) in (list, tuple):
            for row in row_list:
                workbook['WS'][ws_name].set_row(row, height)
        else:
            workbook['WS'][ws_name].set_row(row_list, height)

    def merge_cell(self, workbook, ws_name, rectangle, style=False, data=''):
        """ Merge cell procedure:
            WS: Worksheet where work
            rectangle: list for 2 corners xy data: [0, 0, 10, 5]
            style: setup format for cells
        """
        rectangle.append(data)
        if style:
            rectangle.append(style)
        workbook['WS'][ws_name].merge_range(*rectangle)

    def autofilter(self, workbook, ws_name, rectangle):
        """ Auto filter management
        """
        workbook['WS'][ws_name].autofilter(*rectangle)

    def freeze_panes(self, workbook, ws_name, row, col):
        """ Lock row or column
        """
        workbook['WS'][ws_name].freeze_panes(row, col)

    # -------------------------------------------------------------------------
    # Image management:
    # -------------------------------------------------------------------------
    def clean_odoo_binary(self, odoo_binary_field):
        """ Prepare image data from ODOO binary field:
        """
        return io.BytesIO(base64.decodestring(odoo_binary_field))

    def write_formula(
            self, workbook, ws_name, row, col, formula,
            # format_code,
            value
            ):
        """ Write formula in cell passed
        """
        return workbook['WS'][ws_name].write_formula(
            row, col, formula,
            # self.style[ws_name][format_code],
            # value=value,
            )

    def write_image(
            self, workbook, ws_name, row, col,
            x_offset=0, y_offset=0, x_scale=1, y_scale=1, positioning=2,
            filename=False, data=False, tip='Product image',  # url=False,
            ):
        """ Insert image in cell with extra parameter
            positioning: 1 move + size, 2 move, 3 nothing
        """
        parameters = {
            'tip': tip,
            'x_scale': x_scale,
            'y_scale': y_scale,
            'x_offset': x_offset,
            'y_offset': y_offset,
            'positioning': positioning,
            # 'url': url,
            }

        if data:
            if not filename:
                filename = 'image1.png'  # needed if data present
            parameters['image_data'] = data

        workbook['WS'][ws_name].insert_image(row, col, filename, parameters)
        return True

    def write_image_field_data(
            self, workbook, ws_name, row, col,
            x_offset=0, y_offset=0, x_scale=1, y_scale=1, positioning=2,
            filename=False, odoo_image=False, tip='Product image',
            # url=False,
            ):
        if not odoo_image:
            return False

        return self.write_image(
            workbook=workbook, ws_name=ws_name, row=row, col=col,
            x_offset=x_offset, y_offset=y_offset,
            x_scale=x_scale, y_scale=y_scale,
            positioning=positioning, filename=filename,
            data=self.clean_odoo_binary(odoo_image),
            tip=tip,  # url=url,
        )

    # -------------------------------------------------------------------------
    # Miscellaneous operations (called directly):
    # -------------------------------------------------------------------------
    def write_total_xls_line(
            self, workbook, ws_name, row, total_columns, style_code=False):
        """ Write total line under correct column position
            (use original write function passing every total cell)
        """
        current_total = self.total[ws_name]
        if not current_total:
            _logger.error('No total line needed!')
            return True

        i = 0
        for col in total_columns:
            self.write_xls_line(
                ws_name, row, [current_total[i]],
                style_code=style_code, col=col)
            i += 1

    def write_xls_line(
            self, workbook, ws_name, row, line, style_code=False, col=0,
            total_columns=False,
            ):
        """ Write line in excel file:
            ws_name: Worksheet name where write line
            row: position where write (in ws)
            line: Row passed is a list of element or tuple (element, format)
            style_code: Code for style (see setup format)
            col: add more column data
            total_columns: Tuple with columns need total

            @return: nothing
        """
        def reach_style(ws_name, record):
            """ Convert style code into style of WB (created when inst.)
            """
            res = []
            i = 0
            for item in record:
                i += 1
                if i % 2 == 0:
                    res.append(workbook['style'][ws_name].get(item))
                else:
                    res.append(item)
            return res

        # ---------------------------------------------------------------------
        # Write line:
        # ---------------------------------------------------------------------
        # Setup total list:
        if total_columns and not workbook['total'][ws_name]:
            workbook['total'][ws_name] = [
                0.0 for item in range(0, len(total_columns))]

        # Write every cell of the list:
        style = workbook['style'][ws_name].get(style_code)
        for record in line:
            if type(record) == bool:
                record = ''
            if type(record) not in (list, tuple):
                # Needed?:
                if style:
                    workbook['WS'][ws_name].write(row, col, record, style)
                else:
                    workbook['WS'][ws_name].write(row, col, record)
            elif len(record) == 2:
                # Normal text, format:
                workbook['WS'][ws_name].write(
                    row, col, *reach_style(ws_name, record))
            else:
                # Rich format TODO
                workbook['WS'][ws_name].write_rich_string(
                    row, col, *reach_style(ws_name, record))
            col += 1

        # ---------------------------------------------------------------------
        # Update total columns if necessary
        # ---------------------------------------------------------------------
        if total_columns:
            total_pos = 0
            for total_col in total_columns:
                value = line[total_col]
                # Extract from list/tuple if present:
                if type(value) in (list, tuple):
                    value = value[0]

                if type(value) in (int, float):
                    workbook['total'][ws_name][total_pos] += value
                    total_pos += 1
                else:
                    _logger.error('Float not present in col %s' % total_col)

        # ---------------------------------------------------------------------
        # Setup row height:
        # ---------------------------------------------------------------------
        # TODO if more than one style?
        row_height = workbook['row_height'].get(style, False)
        if row_height:
            workbook['WS'][ws_name].set_row(row, row_height)
        return True

    def rowcol_to_cell(self, row, col, row_abs=False, col_abs=False):
        """ Return row, col format in "A1" notation
        """
        return xl_rowcol_to_cell(row, col, row_abs=row_abs, col_abs=col_abs)

    def write_comment(
            self, workbook, ws_name, row, col, comment, parameters=None):
        """ Write comment in a cell
        """
        cell = self.rowcol_to_cell(row, col)
        if parameters is None:
            parameters = {
                # author, visible, x_scale, width, y_scale, height, color
                # font_name, font_size, start_cell, start_row, start_col
                # x_offset, y_offset
                }
        if comment:
            workbook['WS'][ws_name].write_comment(cell, comment, parameters)

    # -------------------------------------------------------------------------
    # Return operation:
    # -------------------------------------------------------------------------
    def send_mail_to_group(
            self,
            workbook,
            group_name,
            subject, body, filename,
            # Mail data
            ):
        """ Send mail of current workbook to all partner present in group
            passed
            group_name: use format module_name.group_id
            subject: mail subject
            body: mail body
            filename: name of xlsx attached file
        """
        # Send mail with attachment:

        # Pool used
        group_pool = self.env['res.groups']
        model_pool = self.env['ir.model.data']
        thread_pool = self.env['mail.thread']

        # Close before read file:
        attachments = [(
            filename,
            # Raw data:
            open(workbook['filename'], 'rb').read(),
            )]
        group = group_name.split('.')

        # todo change
        group_id = model_pool.get_object_reference(
            group[0], group[1])[1]
        partner_ids = []
        for user in group_pool.browse(group_id).users:
            partner_ids.append(user.partner_id.id)

        thread_pool.message_post(
            False,
            type='email',
            body=body,
            subject=subject,
            partner_ids=[(6, 0, partner_ids)],
            attachments=attachments,
            )

        # if not closed manually
        self.close_workbook(workbook)

    def save_file_as(self, workbook, destination):
        """ Close workbook and save in another place (passed)
        """
        _logger.warning('Save file as: %s' % destination)
        origin = workbook['filename']
        self.close_workbook(workbook)  # if not closed manually
        shutil.copy(origin, destination)
        return True

    def save_binary_xlsx(self, workbook, binary):
        """ Save binary data passed as file temp (returned)
        """
        # todo used?
        b64_file = base64.decodebytes(binary)
        now = fields.Datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
        filename = '/tmp/file_%s.xlsx' % now
        f = open(filename, 'wb')
        f.write(b64_file)
        f.close()
        return filename

    def return_attachment(self, workbook, name, name_of_file=False):
        """ Return attachment passed
            name: Name for the attachment
            name_of_file: file name downloaded
        """
        if not name_of_file:
            now = fields.Datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
            name_of_file = 'report_%s.xlsx' % now

        _logger.info('Return Excel file: %s' % workbook['filename'])

        # todo is necessary?
        temp_id = self.create({
            'fullname': workbook['filename'],
            'b64_file': self.get_b64_from_filename(workbook)
            }).id

        self.close_workbook(workbook)  # if not closed manually
        return {
            'type': 'ir.actions.act_url',
            'name': name,
            'url': '/web/content/excel.report/%s/b64_file/%s?download=true' % (
                temp_id,
                name_of_file,
                ),
            # 'target': 'self',  # XXX Lock button!!!
            }

    # -------------------------------------------------------------------------
    # New Ideas to be implemented:
    # -------------------------------------------------------------------------
    """
    workbook.set_properties({
        'title': 'This is an example spreadsheet',
        'subject': 'With document properties',
        'author': 'Nicola Riolini',
        'manager': 'Nicola Riolini',
        'company': 'Micronaet S.r.l.',
        'category': 'Example spreadsheets',
        'keywords': 'Sample, Example, Properties',
        'created':  datetime.date(2018, 1, 1),
        'comments': 'Created with Python and XlsxWriter'})

    worksheet.repeat_rows()

    worksheet.fit_to_pages()

    set_print_scale()

    worksheet.set_h_pagebreaks()
    worksheet.set_v_pagebreaks()
    """
