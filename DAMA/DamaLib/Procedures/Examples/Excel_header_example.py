import logging

from DamaLib.common.utils.color import Color
from DamaLib.common.excel.sheetformat.Cells import cell as xl_cells, Police, Alignements, Borders,Fill

log = logging.getLogger(__name__)

class style_1(object):

    def __init__(self, worksheet) -> None:
        
        self.ws = worksheet

    def freeze(self, cell:str) -> None:
        xl_cells(self.ws, [cell]).freeze()
        log.info('Freeze %s Done', cell)

    def header_1(self, start_cell:str, end_cell:str) -> None:
        #Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])
        
        wc.rows_height(14)
        wc.police(Police(name='arial', size=14, color=Color.BLACK, bold=True))
        wc.alignment(Alignements(horizontal='center', vertical='center', wrap_text=True))
        wc.color_fill(Fill(fgColor=Color.ORANGE))
        border_h1 = Borders(
            sides= ('left','right','top','bottom'),
            border_style= 'thin',
            color= Color.BLACK)
        wc.borders(border_h1)

        wc.borders(Borders(
            sides=('rangeOutline'),
            border_style= 'thick',
            color= Color.BLACK))
        
        log.info('Header 1 Done (%s:%s)', wc.start_cell.name,wc.end_cell.name)

    def header_2(self, start_cell:str, end_cell:str) -> None:
        #Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.police(Police(size=11, color= Color.BLACK))
        wc.alignment(Alignements(horizontal= 'center', vertical='center', wrap_text=True))
        wc.color_fill(Fill(fgColor= Color.YELLOW))
        wc.borders(Borders(
            sides=('left','right','top','bottom'),
            border_style= 'thin',
            color=Color.BLACK))

        log.info('Header 2 Done (%s:%s)',wc.start_cell.name,wc.end_cell.name)

    def normal (self, start_cell:str, end_cell:str) -> None:
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.police(Police(size=11, color= 'AAA000'))
        wc.alignment(Alignements(horizontal= 'center', vertical='center'))
        wc.color_fill(Fill(fgColor=Color.WHITE))
        wc.borders(Borders(
            sides=('left','right','top','bottom'),
            border_style= 'thin',
            color=Color.BLACK))

    def scientific_number (self, start_cell:str, end_cell:str) -> None:
        self.normal(start_cell, end_cell)
        
        #Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])
        wc.number_format('#.00##E+00')

        log.info('Scientific number Done (%s:%s)',wc.start_cell.name,wc.end_cell.name)

    def separator (self, start_cell:str, end_cell:str) -> None:
        #Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.merge_cells()
        wc.color_fill(Fill(fgColor= Color.BLACK))
        wc.cols_width(2)

        log.info('Separator Done (%s:%s)',wc.start_cell.name,wc.end_cell.name)