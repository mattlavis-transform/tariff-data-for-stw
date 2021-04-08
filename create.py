from classes.measure_type_parser import MeasureTypeParser
from classes.measure_parser import MeasureParser
from classes.prohibition_parser import ProhibitionParser
from classes.document_parser import DocumentParser
from classes.complex_measure_parser import ComplexMeasureParser
import classes.globals as g

g.excel.create_excel()

parser = MeasureTypeParser()
parser.query()
parser.write()

parser = MeasureParser()
parser.query()
parser.write()

parser = ProhibitionParser()
parser.query()
parser.write()

parser = DocumentParser()
parser.query()
parser.write()

parser = ComplexMeasureParser()
parser.query()
parser.write()

g.excel.close_excel()
