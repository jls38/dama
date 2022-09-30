import os
import logging

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.excel.workbook.xlfile import xlfile

from Procedures.Excel_header_example import style_1 as Header_1
from Procedures.Excel_chart_example import style_1 as Chart_1

log = logging.getLogger(__name__)

class workbook_format_1(object):

    def __init__(self, filepath:str='') -> None:
        if filepath == '':
            open_properties(default_extension='.xlsx', filetypes=(('Excel file (.xlsx)', '.xlsx'),('All files', '*,*')))
            filepath = GetPathFromDialog().openfile('Open file')

        self._path = filepath
        self.wb = xlfile.load_workbook(self._path)
        self.ws = None

    def apply(self) -> None:
        log.info('Apply template to file %s: Start', os.path.basename(self._path))
        #Apply template to sheets
        self.sheet1_template(sheet='Feuille1')
        
        log.info('Template to file %s: Done', os.path.basename(self._path))

    def sheet1_template (self, sheet:str) -> None:
        self.ws = self.wb[sheet]
        H1 = Header_1(self.ws)
        #General
        H1.freeze('B2')
        #Titles
        H1.header_1('A1', '1')
        #Title 2
        H1.header_2('A2', 'A')
        #Output Data
        H1.scientific_number('B2', 'B')
        #Dissociate
        H1.separator('C1', 'C10')
        #Charte1
        chart_1 = Chart_1.create_ScatterChart(self.ws, 'K2')
        Chart_1.add_serie_to_ScatterChart(self.ws, chart_1, self.ws, 'A1', 'A10', self.ws, 'B1', 'B10')
        chart_2 = Chart_1.create_ScatterChart(self.ws, 'K7')
        Chart_1.add_serie_to_ScatterChart(self.ws, chart_2, self.ws, 'D1', 'D', self.ws, 'E1', 'E')
        Chart_1.fuse2charts(self.ws, chart_1, chart_2, 'K20')
        #Save template
        self.wb.save(self._path)
        log.info('Template sheet %s : Done', sheet)