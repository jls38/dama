from ast import List
import logging
from copy import copy
from dataclasses import dataclass
from typing import Optional

import openpyxl as opxl
from openpyxl import utils as opxl_utils, styles as opxl_styles
from openpyxl.worksheet.worksheet import Worksheet as OpxlWorkSheet

from DamaLib.common.excel.sheetformat.Coordinates import CellsCoordinates, DMSelect
from DamaLib.common.utils.color import ColorConverter as ColorConverter
from DamaLib.common.decorators.apply_to_dec import for_all_methods
from DamaLib.common.decorators.check import check_method_input
from DamaLib.common.decorators.debug import debug_method, debug_class

log = logging.getLogger(__name__)

@for_all_methods(check_method_input(("",)), "")
class Sheet(CellsCoordinates):
    def __init__(self, select:DMSelect) -> None:
        super().__init__()
        self.selected_datas = self.set_selection(select)
        self.worksheet = self.selected_datas.worksheet
        self.cell = self.selected_datas.worksheet.cell
        self.rows_range = self.selected_datas.rows_range
        self.cols_range = self.selected_datas.columns_range

    def copy_template_sheet(self, template_path: str, template_sheet: str):
        """Copy the style of cells from a sheet"""
        template_workbook = opxl.load_workbook(template_path)
        template_worksheet = template_workbook[template_sheet]

        for row in template_worksheet.rows:
            for cell in row:
                new_cell = self.cell(row=cell.row, column=cell.col_idx)
                if cell.has_style:
                    new_cell.value = copy(cell.value)
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
