import xlsxwriter
from datetime import date
from enum import Enum

from xlsxwriter import worksheet

class ExcelSheet():
    """
    Excel sheet helper
    """
    PLAIN = {'num_format': '@'}
    BOLD = {'bold': True}
    DECIMAL = {'num_format': '0.00'}
    NUMBER = {'num_format': '0'}
    CURRENCY = {'num_format': '$#,##0.00'}
    PERCENT = {'num_format': '%#,##0'}

    # Start values at row 1 since Heading will have row 0
    row = {} 
    has_header_row = {}
    worksheets = {}

    def __init__(self, name):
        self.name = name
        today = date.today()

        file_name = f'{name}-report-{today}.xlsx'
        workbook = xlsxwriter.Workbook(file_name)
        self.blue_heading = workbook.add_format({ 'font_color': '#ffffff', 'bg_color': '#0080ff', 'valign': 'vcenter', 'border': 1, 'font_size': 13 }) 
        self.gray_heading = workbook.add_format({ 'font_color': '#ffffff', 'bg_color': '#969696', 'valign': 'vcenter', 'border': 1, 'font_size': 13 })
        self.generic_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13 })
        self.green_text_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13, 'font_color': '#037d50' })
        self.red_text_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13, 'font_color': '#cc0000' })
        # Add a format. Light red fill with dark red text.
        self.red_text_format = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})
        # Add a format. Green fill with dark green text.
        self.green_text_format = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
        self.workbook = workbook

    def add_worksheet(self, worksheet_name):
        """
        add a new worksheet into a dict
        """
        if worksheet_name not in self.worksheets:
            self.worksheets[worksheet_name] = self.workbook.add_worksheet(worksheet_name)
            self.has_header_row[worksheet_name] = False
            self.row[worksheet_name] = 1


    def add_conditional_format_column(self, worksheet_name, column, type='3_color_scale'):
        """
        add conditional formatting to a column
        """
        # Write a conditional format over a range.
        self.worksheets[worksheet_name].conditional_format(1, column, self.row[worksheet_name], column, {'type': type})


    def add_header_row(self, worksheet_name, headers):
        """
        Add headers at the top of sheet
        Only write header once for this worksheet
        """
        if self.has_header_row[worksheet_name]:
            return
        ws = self.worksheets[worksheet_name]
        
        for col_num, header in enumerate(headers):
            ws.write(0, col_num, header, self.blue_heading)
        self.has_header_row[worksheet_name] = True


    def add_autofilter(self, worksheet_name, last_column):
        """
        Enables filtering and sorting
        """
        self.worksheets[worksheet_name].autofilter(0, 0, self.row[worksheet_name], last_column)


    def add_row(self, worksheet_name, values, formats):
        """
        Add values to a new row in worksheet
        """
        ws = self.worksheets[worksheet_name]
        for col_num, value in enumerate(values):
            number_format = self.workbook.add_format(formats[col_num])
            ws.write(self.row[worksheet_name], col_num, value, number_format)
        self.row[worksheet_name] += 1
    
    def close(self):
        """
        Close the workbook
        """
        self.workbook.close()