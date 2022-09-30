import logging

from DamaLib.common.utils.color import Color
from DamaLib.common.excel.sheetformat.Coordinates import DMSelect
from DamaLib.common.excel.sheetformat.sheet import Sheet as XlSheet
from DamaLib.common.excel.sheetformat.Cells import (
    Cell as XlCells,
    FontParameters,
    AlignementParameters,
    BordersParameters,
    FillParameters,
)

log = logging.getLogger(__name__)


class style_1(object):
    def __init__(self, worksheet) -> None:
        self.ws = DMSelect(worksheet=worksheet, start_cell='A0')

    def freeze(self, cell: str) -> None:
        self.ws.start_cell = cell
        wc = XlCells(self.ws)
        wc.freeze_cell()
        log.info("Freeze %s Done", cell)

    def header_1(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        wc.rows_height_param = 14
        wc.fontparam = FontParameters(name="arial", size=14, color=Color.BLACK, bold=True)
        wc.alignparam = AlignementParameters(horizontal="center", vertical="center", wrap_text=True)
        wc.fillparam = FillParameters(fgColor=Color.ORANGE)
        wc.borderparam = BordersParameters(
            sides=("left", "right", "top", "bottom"),
            border_style="thin",
            color=Color.BLACK,
        )

        wc.apply_rows_height()
        wc.apply_font()
        wc.apply_alignment()
        wc.fill_cell()
        wc.apply_borders()

        log.info("Header 1 Done (%s:%s)", wc.selected_datas.start_cell.coordinate, wc.selected_datas.end_cell.coordinate)

    def header_2(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        wc.fontparam = FontParameters(size=11, color=Color.BLACK)
        wc.alignparam = AlignementParameters(horizontal="center", vertical="center", wrap_text=True)
        wc.fillparam = FillParameters(fgColor=Color.YELLOW)
        wc.borderparam = BordersParameters(
                sides=("left", "right", "top", "bottom"),
                border_style="thin",
                color=Color.BLACK,
            )
        
        wc.apply_font()
        wc.apply_alignment()
        wc.fill_cell()
        wc.apply_borders()

        log.info("Header 2 Done (%s:%s)", wc.selected_datas.start_cell.coordinate, wc.selected_datas.end_cell.coordinate)

    def normal(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        wc.fontparam = FontParameters(size=11, color="AAA000")
        wc.alignparam = AlignementParameters(horizontal="center", vertical="center")
        wc.fillparam = FillParameters(fgColor=Color.WHITE)
        wc.borderparam = BordersParameters(
            sides=("left", "right", "top", "bottom"),
            border_style="thin",
            color=Color.BLACK,
        )
        
        wc.apply_font()
        wc.apply_alignment()
        wc.fill_cell()
        wc.apply_borders()

    def scientific_number(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        self.normal(start_cell, end_cell)
        wc.numformatparam = "#.00##E+00"
        wc.apply_number_format()

        log.info("Scientific number Done (%s:%s)", wc.selected_datas.start_cell.coordinate, wc.selected_datas.end_cell.coordinate)

    def separator(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        wc.merge_cells()
        wc.fillparam = FillParameters(fgColor=Color.BLACK)
        wc.cols_width_param = 2
        wc.fill_cell()
        wc.apply_cols_width()

        log.info("Separator Done (%s:%s)", wc.selected_datas.start_cell.coordinate, wc.selected_datas.end_cell.coordinate)

    def outline(self, start_cell: str, end_cell: str) -> None:
        self.ws.start_cell = start_cell
        self.ws.end_cell = end_cell
        wc = XlCells(self.ws)

        wc.borderparam = BordersParameters(sides=("rangeOutline"), border_style="thick", color=Color.BLACK)
        wc.apply_borders()

        log.info("Outline Done (%s:%s)", wc.selected_datas.start_cell.coordinate, wc.selected_datas.end_cell.coordinate)

