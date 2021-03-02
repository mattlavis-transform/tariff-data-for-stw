import os
import xlsxwriter
from dotenv import load_dotenv

from classes.database import Database
import classes.globals as g


class MeasureTypeParser(object):
    def __init__(self):
        load_dotenv('.env')
        self.database_url = os.getenv('DATABASE_UK')

        self.sql = """
        select mt.measure_type_series_id, mt.measure_type_id, 
        mtd.description, count(m.*)
        from utils.materialized_measures_real_end_dates m, measure_types mt, measure_type_descriptions mtd 
        where mt.measure_type_id = mtd.measure_type_id
        and m.measure_type_id = mt.measure_type_id 
        and mt.validity_end_date is null
        and mt.measure_type_series_id in ('A', 'B')
        and (m.validity_end_date is null or cast(m.validity_end_date as timestamp) > current_date)
        and mt.trade_movement_code != '1'
        group by mt.measure_type_series_id, mt.measure_type_id, 
        mtd.description
        order by mt.measure_type_id;
        """

    def query(self):
        d = Database()
        rows = d.run_query(self.sql)
        self.measure_types = []
        for row in rows:
            measure_type = MeasureType(row[0], row[1], row[2], row[3])
            self.measure_types.append(measure_type)

    def write(self):
        # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = g.excel.workbook.add_worksheet()
        worksheet.name = "Measure types"

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Write headers
        worksheet.write(0, 0, "Series", g.excel.format_header)
        worksheet.write(0, 1, "Measure type ID", g.excel.format_header)
        worksheet.write(0, 2, "Measure type description", g.excel.format_header)
        worksheet.write(0, 3, "Measure count", g.excel.format_header)
        
        widths = [15, 15, 80, 20]
        for i in range(0, len(widths)):
            worksheet.set_column(i, i, widths[i])
        worksheet.freeze_panes(1, 0)

        # Iterate over the data and write it out row by row.
        for measure_type in self.measure_types:
            worksheet.write(row, 0, measure_type.measure_type_series_id)
            worksheet.write(row, 1, measure_type.measure_type_id)
            worksheet.write(row, 2, measure_type.description)
            worksheet.write(row, 3, measure_type.measure_count)
            row += 1

            
        my_range = 'A1:D' + str(row)
        worksheet.autofilter(my_range)


class MeasureType(object):
    def __init__(self, measure_type_series_id, measure_type_id, description, measure_count):
        self.measure_type_series_id = measure_type_series_id
        self.measure_type_id = measure_type_id
        self.description = description
        self.measure_count = measure_count
