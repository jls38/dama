import logging
import re
from dataclasses import dataclass

from openpyxl import utils as opxl_utils
from openpyxl.utils import cell as opxl_cell
from openpyxl.worksheet.worksheet import Worksheet as opxl_worksheet

log = logging.getLogger(__name__)

@dataclass
class _cell:
    name:str
    col_letter = str()
    col_number = int()
    row_number = int()

    def __post_init__ (self):
        self.col_letter = opxl_cell.coordinate_from_string(self.name)[0]
        self.col_number = int(opxl_cell.column_index_from_string(self.col_letter))
        self.row_number = int(opxl_cell.coordinate_from_string(self.name)[1])

#@DebugClass('__init__')
class cells_coordinates (object):
    def __init__(self, worksheet:opxl_worksheet) -> None:
        #Set worksheet
        if not self._isValidExcelWorksheet(worksheet):
            raise TypeError('Worksheet invalid')
        self.worksheet = worksheet

        self.scell = str
        self.ecell = str
        self.start_cell = _cell
        self.end_cell = _cell
        self.rows_range = range
        self.cols_range = range

    def set_cells (self, cells_range:list):
        """cells_range = [start_cell:str, end_cell:str=start_cell]"""
        for cell in cells_range:
            if type(cell) != str:
                raise TypeError('cells_range type must be list[str]')

        self.scell = cells_range[0]
            
        if len(cells_range) == 1:
            self.ecell = self.scell
        elif len(cells_range) == 2:
            self.ecell = cells_range[1]
        else:
            raise ValueError('range must be on the forme [start_cell, end_cell:optional]')
    
        #Set start cell
        if not self._isValidExcelCell(self.scell):
            raise ValueError('Start cell format error')
        self.start_cell = _cell(self.scell)

        #Set end cell
        if self.ecell == self.scell:
            self.end_cell = self.start_cell
        elif self._isValidExcelCell(self.ecell):
            self.end_cell = _cell(self.ecell)
        elif self._isValidExcelColumn(self.ecell) or self._isValidExcelRow(self.ecell):
            self.end_cell = _cell(self._set_coordinate(self.ecell))
        else:
            raise ValueError('End cell format error')

        #Set up rows and columns ranges
        if not self._isValidCoordinates():
            raise ValueError('Start cell has to be positioned on the top-left selection')
        self._set_ranges()

    def _isValidExcelWorksheet (self, ws:opxl_worksheet) -> bool:
        if type(ws) != opxl_worksheet:
            return False
        return True
        
    def _isValidExcelCell (self, cell:str) -> bool:
        #Cell format up to A1-XDF999999
        if type(cell) != str:
            raise TypeError('Type error')

        m = re.fullmatch(r'^([A-Z]{1,3})([1-9]\d*)$', cell)
        if not m:
            return False
      
        letters = m.group(1)
        numbers = m.group(2)
        if len(letters) == 3 and letters > 'XDF':
            log.error("Cell's column value out of range")
            return False
        elif int(numbers) > 1048576:
            log.error("Cell's row value out of range")
            return False

        return True

    def _set_coordinate(self, c:str) -> str | None:
        """Generate cells range from starting cell to end cells:
        Set cell name from column number or row number: 
            Get the last row from column
            Get the last column from row
        """
        if self._isValidExcelRow(c):
            col = int(self.start_cell.col_number)
            while self.worksheet.cell(row=int(c), column=col).value is not None:
                col +=1
            max_col_letter = opxl_utils.get_column_letter(col -1)
            return f"{max_col_letter}{c}"

        if self._isValidExcelColumn(c):
            row = self.start_cell.row_number
            col = opxl_cell.column_index_from_string(c)
            while self.worksheet.cell(row=row, column=col).value is not None:
                row +=1
            max_row = row - 1
            return f"{c}{max_row}"
        
        return None
        
    def _isValidExcelColumn (self, col:str) -> bool:
        if type(col) != str:
            return False
        if not col.isalpha():
            return False
        if len(col) > 3 :
            return False
        if len(col) == 3 :
            if col > 'XDF':
                return False
        return True

    def _isValidExcelRow (self, row:str) -> bool:
        if type(row) != str:
            return False
        if not row.isnumeric():
            return False
        if int(row) > 1048576:
            return False
        return True

    def _isValidCoordinates (self) -> bool:
        if self.start_cell.col_number > self.end_cell.col_number:
            return False

        if self.start_cell.row_number > self.end_cell.row_number:
            return False

        return True

    def _set_ranges(self) -> range:
        self.rows_range = range(self.start_cell.row_number, self.end_cell.row_number + 1)
        self.cols_range = range(self.start_cell.col_number, self.end_cell.col_number + 1)