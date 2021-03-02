import os
import xlsxwriter

from classes.database import Database
import classes.globals as g


class DocumentParser(object):
    def __init__(self):
        self.sql = """
        with cte_document_codes as (
            SELECT cd1.certificate_type_code::text || cd1.certificate_code::text AS code, cd1.description
            FROM certificate_descriptions cd1, certificates c
            WHERE c.certificate_code::text = cd1.certificate_code::text AND c.certificate_type_code::text = cd1.certificate_type_code::text AND (cd1.oid IN ( SELECT max(cd2.oid) AS max
            FROM certificate_descriptions cd2
            WHERE cd1.certificate_type_code::text = cd2.certificate_type_code::text
            AND cd1.certificate_code::text = cd2.certificate_code::text))
            ORDER BY cd1.certificate_type_code, cd1.certificate_code
        ), cte_measure_conditions as (
            select measure_sid, certificate_type_code || certificate_code as code
            from measure_conditions
            where certificate_type_code is not null
        )
        select m.measure_sid, m.goods_nomenclature_item_id, m.geographical_area_id,
        m.measure_type_id, dc.code, dc.description
        from utils.materialized_measures_real_end_dates m,
        cte_document_codes dc, cte_measure_conditions as mc, measure_types mt
        where m.measure_sid = mc.measure_sid
        and m.measure_type_id = mt.measure_type_id 
        and mt.measure_type_series_id in ('A', 'B')
        and (m.validity_end_date is null or cast(m.validity_end_date as timestamp) > current_date)
        and mt.trade_movement_code != '1'
        and mc.code = dc.code
        order by dc.code, m.measure_type_id, m.goods_nomenclature_item_id;
        """

    def query(self):
        d = Database()
        rows = d.run_query(self.sql)
        self.document_codes = []
        for row in rows:
            document_code = Document(row[0], row[1], row[2], row[3], row[4], row[5])
            self.document_codes.append(document_code)

    def write(self):
        # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = g.excel.workbook.add_worksheet()
        worksheet.name = "Document codes"

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Write headers
        worksheet.write(0, 0, "Measure SID", g.excel.format_header)
        worksheet.write(0, 1, "Commodity code", g.excel.format_header)
        worksheet.write(0, 2, "Geography", g.excel.format_header)
        worksheet.write(0, 3, "Measure type ID", g.excel.format_header)
        worksheet.write(0, 4, "Code", g.excel.format_header)
        worksheet.write(0, 5, "Code description", g.excel.format_header)
        
        widths = [15, 20, 20, 20, 15, 100]
        for i in range(0, len(widths)):
            worksheet.set_column(i, i, widths[i])
        worksheet.freeze_panes(1, 0)

        # Iterate over the data and write it out row by row.
        for document_code in self.document_codes:
            worksheet.write(row, 0, document_code.measure_sid)
            worksheet.write(row, 1, document_code.goods_nomenclature_item_id)
            worksheet.write(row, 2, document_code.geographical_area_id)
            worksheet.write(row, 3, document_code.measure_type_id)
            worksheet.write(row, 4, document_code.code)
            worksheet.write(row, 5, document_code.description)
            row += 1
            
        my_range = 'A1:F' + str(row)
        worksheet.autofilter(my_range)


class Document(object):
    def __init__(self, measure_sid, goods_nomenclature_item_id, geographical_area_id, measure_type_id, code, description):
        self.measure_sid = measure_sid
        self.goods_nomenclature_item_id = goods_nomenclature_item_id
        self.geographical_area_id = geographical_area_id
        self.measure_type_id = measure_type_id
        self.description = description
        self.code = code
