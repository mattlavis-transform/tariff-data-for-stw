import os
import xlsxwriter

from classes.database import Database
import classes.globals as g


class MeasureParser(object):
    def __init__(self):
        self.sql = """
        with mc as (
            select measure_sid, certificate_type_code || certificate_code as code
            from measure_conditions
            where certificate_type_code is not null
        )
        select mt.measure_type_id, 
        mtd.description, m.measure_sid, m.goods_nomenclature_item_id,
        m.geographical_area_id,
        string_agg(distinct mc.code, ',' order by mc.code) as codes
        from utils.materialized_measures_real_end_dates m, measure_types mt, measure_type_descriptions mtd, mc
        where mt.measure_type_id = mtd.measure_type_id
        and m.measure_sid = mc.measure_sid 
        and m.measure_type_id = mt.measure_type_id 
        and mt.validity_end_date is null
        and mt.measure_type_series_id in ('A', 'B')
        and (m.validity_end_date is null or cast(m.validity_end_date as timestamp) > current_date)
        and mt.trade_movement_code != '1'
        group by mt.measure_type_id, 
        mtd.description, m.measure_sid , m.goods_nomenclature_item_id,
        m.geographical_area_id
        order by mt.measure_type_id, m.goods_nomenclature_item_id 
        """

    def query(self):
        d = Database()
        rows = d.run_query(self.sql)
        self.measure_types = []
        for row in rows:
            measure_type = Measure(row[0], row[1], row[2], row[3], row[4], row[5])
            self.measure_types.append(measure_type)

    def write(self):
        # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = g.excel.workbook.add_worksheet()
        worksheet.name = "Measures"

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Write headers
        worksheet.write(0, 0, "Measure type ID", g.excel.format_header)
        worksheet.write(0, 1, "Measure type description", g.excel.format_header)
        worksheet.write(0, 2, "Measure SID", g.excel.format_header)
        worksheet.write(0, 3, "Commodity code", g.excel.format_header)
        worksheet.write(0, 4, "Geography", g.excel.format_header)
        worksheet.write(0, 5, "Codes", g.excel.format_header)
        
        widths = [15, 80, 20, 20, 20, 50]
        for i in range(0, len(widths)):
            worksheet.set_column(i, i, widths[i])
        worksheet.freeze_panes(1, 0)

        # Iterate over the data and write it out row by row.
        for measure_type in self.measure_types:
            worksheet.write(row, 0, measure_type.measure_type_id)
            worksheet.write(row, 1, measure_type.description)
            worksheet.write(row, 2, measure_type.measure_sid)
            worksheet.write(row, 3, measure_type.goods_nomenclature_item_id)
            worksheet.write(row, 4, measure_type.geographical_area_id)
            worksheet.write(row, 5, measure_type.codes)
            row += 1
            
        my_range = 'A1:F' + str(row)
        worksheet.autofilter(my_range)


class Measure(object):
    def __init__(self, measure_type_id, description, measure_sid, goods_nomenclature_item_id, geographical_area_id, codes):
        self.measure_type_id = measure_type_id
        self.description = description
        self.measure_sid = measure_sid
        self.goods_nomenclature_item_id = goods_nomenclature_item_id
        self.geographical_area_id = geographical_area_id
        self.codes = codes
