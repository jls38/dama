from ast import List
import logging
import re
from dataclasses import dataclass
from typing import Optional

from openpyxl import utils as opxl_utils
from openpyxl.utils import cell as opxl_cell
from openpyxl.worksheet.worksheet import Worksheet as opxl_worksheet

from DAMA.constants.constant import MAX_COLUMN_EXL, MAX_ROW_EXL

log = logging.getLogger(__name__)


@dataclass
class _Cell:
    name: str
    col_letter = str()
    col_number = int()
    row_number = int()

    def __post_init__(self):
        self.col_letter = opxl_cell.coordinate_from_string(self.name)[0]
        self.col_number = int(opxl_cell.column_index_from_string(self.col_letter))
        self.row_number = int(opxl_cell.coordinate_from_string(self.name)[1])


# @DebugClass('__init__')
class CellsCoordinates(object):
    def __init__(self, worksheet: opxl_worksheet) -> None:
        # Set worksheet
        if not CellsCoordinates._isValidExcelWorksheet(worksheet):
            raise TypeError("Worksheet invalid")
        self.worksheet = worksheet

        self.scell = str
        self.ecell = str
        self.start_cell = _Cell
        self.end_cell = _Cell
        self.rows_range = range
        self.cols_range = range

    def set_cells(self, cells_range: List[str]):
        """cells_range = [start_cell:str, end_cell:str=start_cell]"""
        for cell in cells_range:
            if not type(cell) == str:
                raise TypeError("cells_range type must be list[str]")

        self.scell = cells_range[0]

        match len(cells_range):
            case 1:
                self.ecell = self.scell
            case 2:
                self.ecell = cells_range[1]
            case _:
                raise ValueError(
                    "range must be on the forme [start_cell, end_cell:optional]"
                )

        # Set start cell
        if not CellsCoordinates._isValidExcelCell(self.scell):
            raise ValueError("Start cell format error")
        self.start_cell = _Cell(self.scell)

        # Set end cell
        if self.ecell == self.scell:
            self.end_cell = self.start_cell
        elif CellsCoordinates._isValidExcelCell(self.ecell):
            self.end_cell = _Cell(self.ecell)
        elif CellsCoordinates._isValidExcelColumn(
            self.ecell
        ) or CellsCoordinates._isValidExcelRow(self.ecell):
            self.end_cell = _Cell(self._set_coordinate(self.ecell))
        else:
            raise ValueError("End cell format error")

        # Set up rows and columns ranges
        if not self._isValidCoordinates():
            raise ValueError(
                "Start cell has to be positioned on the top-left selection"
            )
        self._set_ranges()

    @staticmethod
    def _isValidExcelWorksheet(self, ws: opxl_worksheet) -> bool:
        return type(ws) == opxl_worksheet

    @staticmethod
    def _isValidExcelCell(cell: str) -> bool:
        # Cell format up to A1-XDF999999
        if not type(cell) == str:
            raise TypeError("Type error")

        m = re.fullmatch(r"^([A-Z]{1,3})([1-9]\d*)$", cell)
        if not m:
            return False

        letters = m.group(1)
        numbers = m.group(2)
        if len(letters) == 3 and letters > "XDF":
            log.error("Cell's column value out of range")
            return False
        elif int(numbers) > 1048576:
            log.error("Cell's row value out of range")
            return False

        return True

    def _set_coordinate(self, c: str) -> Optional[str]:
        """Generate cells range from starting cell to end cells:
        Set cell name from column number or row number:
            Get the last row from column
            Get the last column from row
        """
        if CellsCoordinates._isValidExcelRow(c):
            col = int(self.start_cell.col_number)
            while self.worksheet.cell(row=int(c), column=col).value is not None:
                col += 1
            max_col_letter = opxl_utils.get_column_letter(col - 1)
            return f"{max_col_letter}{c}"

        if CellsCoordinates._isValidExcelColumn(c):
            row = self.start_cell.row_number
            col = opxl_cell.column_index_from_string(c)
            while self.worksheet.cell(row=row, column=col).value is not None:
                row += 1
            max_row = row - 1
            return f"{c}{max_row}"

        return None

    @staticmethod
    def _isValidExcelColumn(col: str) -> bool:
        return not (
            not type(col) == str
            or not col.isalpha()
            or len(col) > 3
            or (len(col) == 3 and col > MAX_COLUMN_EXL)
        )
        """
        if not type(col) == str:
            return False
        if not col.isalpha():
            return False
        if len(col) > 3:
            return False
        if len(col) == 3:
            if col > "XDF":
                return False
        return True
        """

    @staticmethod
    def _isValidExcelRow(row: str) -> bool:
        return not (type(row) == str or not row.isnumeric() or int(row) > MAX_ROW_EXL)

        """
        if type(row) != str:
            return False
        if not row.isnumeric():
            return False
        if int(row) > 1048576:
            return False
        return True
        """

    def _isValidCoordinates(self) -> bool:
        return not (
            self.start_cell.col_number > self.end_cell.col_number
            or self.start_cell.row_number > self.end_cell.row_number
        )
        """
        if self.start_cell.col_number > self.end_cell.col_number:
            return False

        if self.start_cell.row_number > self.end_cell.row_number:
            return False

        return True
        """

    def _set_ranges(self) -> range:
        self.rows_range = range(
            self.start_cell.row_number, self.end_cell.row_number + 1
        )
        self.cols_range = range(
            self.start_cell.col_number, self.end_cell.col_number + 1
        )
