import os
from datetime import datetime
import xlsxwriter


class Excel(object):
    def __init__(self):
        pass
    
    def get_filename(self):
        date_string = datetime.now().strftime('%Y-%m-%d')
        self.filename = "uk_tariff_document_codes_" + date_string + ".xlsx"
            
    def get_path(self):
        self.path = os.getcwd()
        self.path = os.path.join(self.path, "output")
        os.makedirs(self.path, exist_ok = True)

    def create_excel(self):
        self.get_path()
        self.get_filename()
        self.excel_filename = os.path.join(self.path, self.filename)
        
        # Open the workbook
        self.workbook = xlsxwriter.Workbook(self.excel_filename)
        
        # Create the formats that will be used in all sheets
        self.format_header = self.workbook.add_format({'bold': True})
        self.format_header.set_align('top')
        self.format_header.set_align('left')
        self.format_header.set_bg_color("#000000")
        self.format_header.set_color("#ffffff")

        self.format_wrap = self.workbook.add_format({'text_wrap': True})
        self.format_wrap.set_align('top')
        self.format_wrap.set_align('left')

        self.format_force_text = self.workbook.add_format({'text_wrap': True})
        self.format_force_text.set_align('top')
        self.format_force_text.set_align('left')
        self.format_force_text.set_num_format('@')
        

    def close_excel(self):
        self.workbook.close()
