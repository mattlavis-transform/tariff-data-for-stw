import os
import xlsxwriter

from classes.database import Database
import classes.globals as g


class ComplexMeasureParser(object):
    def __init__(self):
        self.sql = """
        with cte2 as (
            with cte as (
                select mc.measure_sid, m.goods_nomenclature_item_id, m.geographical_area_id,
                m.measure_type_id, mc.condition_code,
                string_agg(coalesce(mc.certificate_type_code, '') || coalesce(mc.certificate_code, ''), '|') as codes
                from measure_conditions mc, utils.materialized_measures_real_end_dates m, measure_types mt 
                where m.measure_sid = mc.measure_sid 
                and m.measure_type_id = mt.measure_type_id 
                and m.validity_end_date is null
                and mt.measure_type_series_id = 'B'
                and mt.trade_movement_code != '1'
                and m.measure_type_id not in (
                    '464',
                    '474',
                    '495',
                    '494',
                    '496',
                    '492',
                    '491',
                    '493',
                    '484',
                    '485',
                    '481',
                    '482',
                    '483'
                )
                and mc.action_code >= '24'
                group by mc.measure_sid, m.goods_nomenclature_item_id, m.geographical_area_id, m.measure_type_id, mc.condition_code
            )
            select measure_sid, goods_nomenclature_item_id, geographical_area_id, measure_type_id, count(*) as condition_code_group_count
            from cte
            group by measure_sid, goods_nomenclature_item_id, geographical_area_id, measure_type_id
            order by 5 desc, measure_type_id, geographical_area_id, goods_nomenclature_item_id
            )
            select * from cte2 where condition_code_group_count > 1
        """

    def query(self):
        d = Database()
        rows = d.run_query(self.sql)
        self.document_codes = []
        for row in rows:
            document_code = Document(row[0], row[1], row[2], row[3], row[4])
            self.document_codes.append(document_code)

    def write(self):
        # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = g.excel.workbook.add_worksheet()
        worksheet.name = "Complex measures"

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Write headers
        worksheet.write(0, 0, "Measure SID", g.excel.format_header)
        worksheet.write(0, 1, "Commodity code", g.excel.format_header)
        worksheet.write(0, 2, "Geography", g.excel.format_header)
        worksheet.write(0, 3, "Measure type ID", g.excel.format_header)
        worksheet.write(0, 4, "Condition code group count", g.excel.format_header)
        
        widths = [15, 20, 20, 20, 15]
        for i in range(0, len(widths)):
            worksheet.set_column(i, i, widths[i])
        worksheet.freeze_panes(1, 0)

        # Iterate over the data and write it out row by row.
        for document_code in self.document_codes:
            worksheet.write(row, 0, document_code.measure_sid)
            worksheet.write(row, 1, document_code.goods_nomenclature_item_id)
            worksheet.write(row, 2, document_code.geographical_area_id)
            worksheet.write(row, 3, document_code.measure_type_id)
            worksheet.write(row, 4, document_code.condition_code_group_count)
            row += 1
            
        my_range = 'A1:E' + str(row)
        worksheet.autofilter(my_range)


class Document(object):
    def __init__(self, measure_sid, goods_nomenclature_item_id, geographical_area_id, measure_type_id, condition_code_group_count):
        self.measure_sid = measure_sid
        self.goods_nomenclature_item_id = goods_nomenclature_item_id
        self.geographical_area_id = geographical_area_id
        self.measure_type_id = measure_type_id
        self.condition_code_group_count = condition_code_group_count
