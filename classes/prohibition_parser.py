import os
import xlsxwriter

from classes.database import Database
import classes.globals as g


class ProhibitionParser(object):
    def __init__(self):
        self.sql = """
        select mt.measure_type_id, 
        mtd.description, m.measure_sid, m.goods_nomenclature_item_id,
        m.geographical_area_id --, mc.certificate_type_code 
        from measure_types mt, measure_type_descriptions mtd, utils.materialized_measures_real_end_dates m
        left join measure_conditions mc on m.measure_sid = mc.measure_sid 
        where mt.measure_type_id = mtd.measure_type_id
        and mc.measure_sid is null
        and m.measure_type_id = mt.measure_type_id 
        and mt.validity_end_date is null
        and mt.measure_type_series_id in ('A', 'B')
        and (m.validity_end_date is null or cast(m.validity_end_date as timestamp) > current_date)
        and mt.trade_movement_code != '1'
        order by mt.measure_type_id, m.goods_nomenclature_item_id;
        """

    def query(self):
        d = Database()
        rows = d.run_query(self.sql)
        self.measures = []
        for row in rows:
            measure = Measure(row[0], row[1], row[2], row[3], row[4])
            self.measures.append(measure)

    def write(self):
        # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = g.excel.workbook.add_worksheet()
        worksheet.name = "Measures - no conditions"

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Write headers
        worksheet.write(0, 0, "Measure type ID", g.excel.format_header)
        worksheet.write(0, 1, "Measure type description", g.excel.format_header)
        worksheet.write(0, 2, "Measure SID", g.excel.format_header)
        worksheet.write(0, 3, "Commodity code", g.excel.format_header)
        worksheet.write(0, 4, "Geography", g.excel.format_header)
        
        widths = [15, 80, 20, 20, 20]
        for i in range(0, len(widths)):
            worksheet.set_column(i, i, widths[i])
        worksheet.freeze_panes(1, 0)

        # Iterate over the data and write it out row by row.
        for measure in self.measures:
            worksheet.write(row, 0, measure.measure_type_id)
            worksheet.write(row, 1, measure.description)
            worksheet.write(row, 2, measure.measure_sid)
            worksheet.write(row, 3, measure.goods_nomenclature_item_id)
            worksheet.write(row, 4, measure.geographical_area_id)
            row += 1
            
        my_range = 'A1:E' + str(row)
        worksheet.autofilter(my_range)


class Measure(object):
    def __init__(self, measure_type_id, description, measure_sid, goods_nomenclature_item_id, geographical_area_id):
        self.measure_type_id = measure_type_id
        self.description = description
        self.measure_sid = measure_sid
        self.goods_nomenclature_item_id = goods_nomenclature_item_id
        self.geographical_area_id = geographical_area_id