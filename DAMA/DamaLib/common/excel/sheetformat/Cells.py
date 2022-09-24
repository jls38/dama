from ast import List
from distutils.log import debug
import logging
from copy import copy
from dataclasses import dataclass
from typing import Optional

import openpyxl as opxl
from openpyxl import utils as opxl_utils, styles as opxl_styles
from openpyxl.worksheet.worksheet import Worksheet as OpxlWorkSeet

from DamaLib.common.excel.sheetformat.Coordinates import CellsCoordinates
from DamaLib.common.utils.color import color_converter as cmn_color_conv
from DamaLib.common.decorators.apply_to_dec import for_all_methods
from DamaLib.common.decorators.check import check_method_input
from DamaLib.common.decorators.debug import DebugMethod, DebugClass

log = logging.getLogger(__name__)


@dataclass
class Font:
    """
    'underline' must be one of {'none','single', 'double'}
    """

    name: str = None
    size: int | float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: str = "none"
    strike: bool | None = None
    color: str | None = None

    def __post_init__(self):
        if self.color != None:
            self.color = cmn_color_conv(self.color).hexaCode()

        if not self.underline in ("none", "single", "double"):
            raise ValueError("underline must be one of: 'none', 'single', 'double'")


@dataclass
class Borders:
    """
    sides:
        - left | right | top | bottom | diagonal
        - vertical | horizontal
        - outline (for no color)
        - rangeOutline (for range Outline)

    border_style must be one of ['slantDashDot', 'thin', 'double', 'hair', 'dashDot', 'dashDotDot', 'mediumDashDot', 'mediumDashDotDot', 'dotted', 'mediumDashed', 'medium', 'dashed', 'thick']

    color: str (name or hexacode(#))
    """

    sides: tuple
    border_style: str = "none"
    color: Optional[str] = None

    def __post_init__(self):
        if not self.sides == "rangeOutline":
            sides_list = (
                "left",
                "right",
                "top",
                "bottom",
                "diagonal",
                "vertical",
                "horizontal",
                "outline",
            )

            for side in self.sides:
                if not side in sides_list:
                    raise ValueError(f"Sides must be on the list {sides_list}")

        border_list = (
            "none",
            "slantDashDot",
            "thin",
            "double",
            "hair",
            "dashDot",
            "dashDotDot",
            "mediumDashDot",
            "mediumDashDotDot",
            "dotted",
            "mediumDashed",
            "medium",
            "dashed",
            "thick",
        )
        if not self.border_style in border_list:
            raise ValueError(f"Sides must be on the {border_list=}")

        if self.color is not None:
            self.color = cmn_color_conv(self.color).hexaCode()


@dataclass
class FillParameters:
    fgColor: str
    fill_type: Optional[str] = "solid"
    patternType: Optional[str] = None

    def __post_init__(self):
        if self.fgColor is not None:
            self.fgColor = cmn_color_conv(self.fgColor).hexaCode()


@dataclass
class Alignements:
    """
    horizontal must be one of : 'centerContinuous', 'distributed', 'right', 'fill', 'left', 'general', 'justify', 'center'
    vertical must be one of : 'bottom', 'justify', 'distributed', 'center', 'top'
    """

    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrap_text: Optional[bool] = None
    shrink_to_fit: Optional[bool] = None
    text_rotation: int = 0
    indent: int = 0

    def __post_init__(self):
        h_alignment_list = (
            "centerContinuous",
            "distributed",
            "right",
            "fill",
            "left",
            "general",
            "justify",
            "center",
        )

        if self.horizontal is not None and not self.horizontal in h_alignment_list:
            raise ValueError(f"horizontal must be None or one of {h_alignment_list=}")

        v_alignment_list = (
            "bottom",
            "justify",
            "distributed",
            "center",
            "top",
        )

        if self.vertical is not None and not self.vertical in v_alignment_list:
            raise ValueError(f"horizontal must be None or one of {v_alignment_list=}")


# @DebugClass('__init__')
@for_all_methods(check_method_input(("",)), "")
class Cell(CellsCoordinates):
    def __init__(
        self, worksheet: OpxlWorkSeet, cells_range: List[str], font: Font = Font()
    ) -> None:
        """cells_range: [start_cell:str, end_cell:str]"""
        super().__init__(worksheet)
        self.set_cells(cells_range)
        self.cell = self.worksheet.cell
        self.font = font

    def freeze(self) -> None:
        self.worksheet.freeze_panes = self.start_cell.name

    def set_cols_width(self, size: int | float) -> None:
        for col in self.cols_range:
            col = opxl_utils.get_column_letter(col)
            self.worksheet.column_dimensions[col].width = size

    def set_rows_height(self, size: int | float) -> None:
        for row in self.rows_range:
            self.worksheet.row_dimensions[row].height = size

    def fill(self, fill: FillParameters) -> None:
        cells_fill = {str(attr): getattr(fill, attr) for attr in fill.__dict__}
        set_fill = opxl_styles.PatternFill(**cells_fill)

        for r in self.rows_range:
            for c in self.cols_range:
                self.cell(row=r, column=c).fill = set_fill

    def set_font(self, font: Font) -> None:
        self.font = font

    def apply_font(self) -> None:
        cells_Font = {
            str(attr): getattr(self.font, attr) for attr in self.font.__dict__
        }

        for r in self.rows_range:
            for c in self.cols_range:
                self.cell(row=r, column=c).font = opxl_styles.Font(**cells_Font)

    def number_format(self, format: str) -> None:
        """
        'format' : str ('general' | '0' | '#,##0.00' | '#,##0.00E+00' | '#,#0.0% | ...)
        """
        for r in self.rows_range:
            for c in self.cols_range:
                self.cell(row=r, column=c).number_format = format

    def borders(self, bord: Borders) -> None:
        cells_Border = {"border_style": bord.border_style, "color": bord.color}

        if bord.sides == "rangeOutline":
            side = opxl_styles.Side(**cells_Border)
            max_row = len(self.rows_range) - 1
            max_col = len(self.cols_range) - 1

            for i, r in enumerate(self.rows_range):
                for j, c in enumerate(self.cols_range):
                    # Initializing borders
                    cell = self.cell(row=r, column=c)
                    border = opxl_styles.Border(
                        left=cell.border.left,
                        right=cell.border.right,
                        top=cell.border.top,
                        bottom=cell.border.bottom,
                    )
                    # Set border on the side of the selection
                    if j == 0:
                        border.left = side
                    if j == max_col:
                        border.right = side
                    if i == 0:
                        border.top = side
                    if i == max_row:
                        border.bottom = side
                    if i == 0 or i == max_row or j == 0 or j == max_col:
                        cell.border = border

        else:
            # Define borders' format
            borders_format = {s: opxl_styles.Side(**cells_Border) for s in bord.sides}
            # Apply border format to cells range
            for r in self.rows_range:
                for c in self.cols_range:
                    self.cell(row=r, column=c).border = opxl_styles.Border(
                        **borders_format
                    )

    def alignment(self, align: Alignements):
        cAlignment = {str(attr): getattr(align, attr) for attr in align.__dict__}

        for r in self.rows_range:
            for c in self.cols_range:
                self.cell(row=r, column=c).alignment = opxl_styles.Alignment(
                    **cAlignment
                )

    def merge_cells(self) -> None:
        # Check if less than 2 cells are not empty
        cells_not_empty = 0
        for r in self.rows_range:
            for c in self.cols_range:
                if self.cell(row=r, column=c).value != None:
                    cells_not_empty += 1
                if cells_not_empty >= 2:
                    log.error("Impossile to merge cells")
                    break
            else:
                continue
            break
        else:
            self.worksheet.merge_cells(f"{self.start_cell.name}:{self.end_cell.name}")

    def copy_template_sheet(self, template_path: str, template_sheet: str):
        """Copy the style of cells from a sheet"""
        template_workbook = opxl.load_workbook(template_path)
        template_worksheet = template_workbook[template_sheet]

        for row in template_worksheet.rows:
            for cell in row:
                new_cell = self.cell(row=cell.row, column=cell.col_idx)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
