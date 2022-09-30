import logging

from DamaLib.common.utils.color import Color
from DamaLib.common.excel.sheetformat.Cells import (
    Cell as xl_cells,
    FontParameters,
    AlignementParameters,
    BordersParameters,
    FillParameters,
)

log = logging.getLogger(__name__)


class style_1(object):
    def __init__(self, worksheet) -> None:

        self.ws = worksheet

    def freeze(self, cell: str) -> None:
        xl_cells(self.ws, [cell]).freeze_cell()
        log.info("Freeze %s Done", cell)

    def header_1(self, start_cell: str, end_cell: str) -> None:
        # Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.apply_rows_height(14)
        wc.set_font(FontParameters(name="arial", size=14, color=Color.BLACK, bold=True))
        wc.apply_alignment(
            AlignementParameters(horizontal="center", vertical="center", wrap_text=True)
        )
        wc.fill(FillParameters(fgColor=Color.ORANGE))
        border_h1 = BordersParameters(
            sides=("left", "right", "top", "bottom"),
            border_style="thin",
            color=Color.BLACK,
        )
        wc.apply_borders(border_h1)

        wc.apply_borders(
            BordersParameters(sides=("rangeOutline"), border_style="thick", color=Color.BLACK)
        )

        log.info("Header 1 Done (%s:%s)", wc.start_cell.coordinate, wc.end_cell.coordinate)

    def header_2(self, start_cell: str, end_cell: str) -> None:
        # Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.set_font(FontParameters(size=11, color=Color.BLACK))
        wc.apply_alignment(
            AlignementParameters(horizontal="center", vertical="center", wrap_text=True)
        )
        wc.fill(FillParameters(fgColor=Color.YELLOW))
        wc.apply_borders(
            BordersParameters(
                sides=("left", "right", "top", "bottom"),
                border_style="thin",
                color=Color.BLACK,
            )
        )

        log.info("Header 2 Done (%s:%s)", wc.start_cell.coordinate, wc.end_cell.coordinate)

    def normal(self, start_cell: str, end_cell: str) -> None:
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.set_font(FontParameters(size=11, color="AAA000"))
        wc.apply_alignment(AlignementParameters(horizontal="center", vertical="center"))
        wc.fill(FillParameters(fgColor=Color.WHITE))
        wc.apply_borders(
            BordersParameters(
                sides=("left", "right", "top", "bottom"),
                border_style="thin",
                color=Color.BLACK,
            )
        )

    def scientific_number(self, start_cell: str, end_cell: str) -> None:
        self.normal(start_cell, end_cell)

        # Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])
        wc.apply_number_format("#.00##E+00")

        log.info("Scientific number Done (%s:%s)", wc.start_cell.coordinate, wc.end_cell.coordinate)

    def separator(self, start_cell: str, end_cell: str) -> None:
        # Set workcells
        wc = xl_cells(self.ws, [start_cell, end_cell])

        wc.merge_cells()
        wc.fill(FillParameters(fgColor=Color.BLACK))
        wc.apply_cols_width(2)

        log.info("Separator Done (%s:%s)", wc.start_cell.coordinate, wc.end_cell.coordinate)
