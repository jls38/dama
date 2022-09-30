import logging
from copy import copy
from dataclasses import dataclass
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet as opxl_worksheet
from openpyxl.chart.scatter_chart import ScatterChart as opxl_ScatterChart
from openpyxl.chart.line_chart import LineChart as opxl_LineChart
from openpyxl.chart.bar_chart import BarChart as opxl_BarChart
from openpyxl.chart.series import XYSeries as opxl_XYSeries
from openpyxl.chart.series_factory import SeriesFactory as opxl_chart_SeriesFactory
from openpyxl.chart.reference import Reference as opxl_chart_Reference
from openpyxl.chart.error_bar import ErrorBars as opxl_ErrorBars
from openpyxl.chart.data_source import (
    NumDataSource as opxl_NumDataSource,
    NumRef as opxl_NumRef,
)
from openpyxl.chart.trendline import Trendline as opxl_Trendline

from DamaLib.common.excel.sheetformat.Coordinates import CellsCoordinates, DMSelect
from DamaLib.common.utils.color import ColorConverter as ColorConverter
from DamaLib.common.decorators.check import check_method_input, check_dataclass_input


log = logging.getLogger(__name__)

@check_dataclass_input
@dataclass
class ChartFormat:
    title: Optional[str] = "Title"
    style: Optional[int|float] = 13
    x_axis_title: Optional[str] = "x_axis title"
    y_axis_title: Optional[str] = "y_axis tile"
    x_axis_majorGridlines: Optional[bool] = None
    y_axis_majorGridlines: Optional[bool] = None

    def __post_init__(self):
        pass    

@check_dataclass_input
@dataclass
class SeriesData:
    X_axis: DMSelect
    Y_axis: DMSelect
    title_from_data: bool = True

    def __post_init__(self):
        pass

@check_dataclass_input
@dataclass
class SerieFormat:
    serie_title: Optional[str] = None
    marker_shape: Optional[str] = None
    marker_size: Optional[int | float] = None
    marker_color: Optional[str] = None
    hide_lines: bool = True

    def __post_init__(self):
        if not self.marker_color == None:
            self.marker_color = ColorConverter(self.marker_color).hexaCode()

@check_dataclass_input
@dataclass
class ErrorData:
    err_Val: Optional[int | float] = None
    cust_Errmin: Optional[DMSelect] = None
    cust_Errmax: Optional[DMSelect] = None

    def __post_init__(self):
        pass

@check_dataclass_input
@dataclass
class ErrorFormat:
    """
    err_Dir: must be one of {'x', 'y'}
    err_Bar_Type: must be one of {'minus','both','plus'}
    err_val_type: must be one of {'fixedVal', 'cust', 'stdDev', 'percentage', 'stdErr'}
    """

    err_Dir: str
    err_Bar_Type: str
    err_val_type: str
    No_endCap: bool = False

    def __post_init__(self):
        errDir =('x', 'y')
        if not self.err_Dir in errDir:
            raise ValueError(f"err_Dir must be one of {errDir=}")

        errBarType = ("minus", "both", "plus")
        if not self.err_Bar_Type in errBarType:
            raise ValueError(f"err_Bar_Type must be one of {errBarType=}")

        errValType = ("fixedVal", "cust", "stdDev","percentage","stdErr")
        if not self.err_val_type in errValType:
            raise ValueError(
                f"err_val_type must be one of {errValType=}"
            )

@check_dataclass_input
@dataclass
class TrendLine:
    """
    trenline_type: must be one of {'poly', 'power', 'exp', 'log', 'movingAvg', 'linear'}
    """

    trendline_Type: str
    disp_Eq: bool = False
    disp_RSqr: bool = False
    intercept: Optional[int | float] = None
    order: Optional[int] = None
    period: Optional[int] = None

    def __post_init__(self):
        trendlineType = ("poly", "power", "exp", "log", "movingAvg", "linear")
        if not self.trendline_Type in trendlineType:
            raise ValueError(
                f"trenline_type: must be one of {trendlineType}"
            )

class Charts(object):
    def __init__(self) -> None:
        self._chartFormat = ChartFormat
        self._serieFormat = SerieFormat
        self._errFormat = ErrorFormat
        self._validechartformat = (opxl_ScatterChart, opxl_BarChart, opxl_LineChart)

    @property
    def chart_format(self):
        return self._chartFormat

    @check_method_input('')
    @chart_format.setter
    def chart_format(self, format:ChartFormat):
        self._chartFormat = format

    @property
    def serie_format(self):
        return self._serieFormat

    @check_method_input('')
    @serie_format.setter
    def serie_format(self, format:SerieFormat):
        self._serieFormat = format

    @property
    def error_format(self):
        return self._errFormat

    @check_method_input('')
    @error_format.setter
    def error_format(self, format: ErrorFormat):
        self._errFormat = format

    @check_method_input('')
    def create_chart(
        self, chart_type: str, worksheet: opxl_worksheet, position: str
    ) -> opxl_ScatterChart | opxl_BarChart | opxl_LineChart:
        """
        chart_type must be one of('LineChart', 'BarChart', 'ScatterChart')
        Position : cell coordinate of chart's top-left corner
        """
        if not CellsCoordinates._isValidExcelWorksheet(worksheet):
            raise TypeError("Worksheet type error")

        if not CellsCoordinates._isValidExcelCell(position):
            raise Exception('Cell coordinate not valid')

        match chart_type:
            case "ScatterChart":
                c = opxl_ScatterChart()
            case "BarChart":
                c = opxl_BarChart()
            case "LineChart":
                c = opxl_LineChart()
            case _:
                raise ValueError("Chart type incorrect or none implemented")

        c.title = self.chart_format.title
        c.style = self.chart_format.style
        c.x_axis.title = self.chart_format.x_axis_title
        c.y_axis.title = self.chart_format.y_axis_title
        c.x_axis.majorGridlines = self.chart_format.x_axis_majorGridlines
        c.y_axis.majorGridlines = self.chart_format.y_axis_majorGridlines

        worksheet.add_chart(c, position)
        return c        

    @check_method_input('')
    @staticmethod
    def add_series_to_chart(
        c: opxl_ScatterChart | opxl_BarChart | opxl_LineChart, SData: SeriesData
    ) -> opxl_XYSeries:

        # Get x values
        xvalues = Charts._setSerieValue(SData.X_axis)

        # Get y values
        yvalues = Charts._setSerieValue(SData.Y_axis)

        series = opxl_chart_SeriesFactory(
            yvalues, xvalues, title_from_data=SData.title_from_data
        )
        c.series.append(series)

    @check_method_input('')
    @staticmethod
    def get_series_from_chart(
        c: opxl_ScatterChart | opxl_BarChart | opxl_LineChart
    ) -> opxl_XYSeries:
        """
        Return series from chart
        """
        return c.series

    @check_method_input('')
    def apply_format_to_serie(self, serie: opxl_XYSeries) -> None:
        """
        serie = chart.series[i]
        """
        serie.title = self.serie_format.serie_title
        serie.marker.symbol = self.serie_format.marker_shape
        serie.marker.size = self.serie_format.marker_size
        serie.marker.graphicalProperties.solidFill = (
            self.serie_format.marker_color
        )  # Marker filling
        serie.marker.graphicalProperties.line.solidFill = (
            self.serie_format.marker_color
        )  # Marker outline
        serie.graphicalProperties.line.noFill = self.serie_format.hide_lines  # hide lines

    @check_method_input('')
    def add_to_serie_error_bars(self, serie: opxl_XYSeries, errData: ErrorData):
        """
        serie: series[i]

        if {'stdErr'} used, no values needs to be specified
        if {'fixedVal', 'stdDev', 'percentage'} used, specified err_Val
        if {'cust'} used, specified cust_Errmin_List and/or cust_Errmax_List
        """
        match self.error_format.err_val_type:
            case 'cust':
                err_Val = None
                match self.error_format.err_Bar_Type:
                    case "minus":
                        Minus = Charts._setSerieValue(errData.cust_Errmin)
                        MinusNumDataSource = opxl_NumDataSource(opxl_NumRef(Minus))
                    case "plus":
                        Plus = Charts._setSerieValue(errData.cust_Errmax)
                        PlusNumeDataSource = opxl_NumDataSource(opxl_NumRef(Plus))
                    case "both":
                        Minus = Charts._setSerieValue(errData.cust_Errmin)
                        MinusNumDataSource = opxl_NumDataSource(opxl_NumRef(Minus))

                        Plus = Charts._setSerieValue(errData.cust_Errmax)
                        PlusNumeDataSource = opxl_NumDataSource(opxl_NumRef(Plus))
            case _:
                MinusNumDataSource = None
                PlusNumeDataSource = None
                err_Val = errData.err_Val

        # Add error to serie
        serie.errBars = opxl_ErrorBars(
            errDir= self.error_format.err_Dir,
            errBarType= self.error_format.err_Bar_Type,
            errValType= self.error_format.err_val_type,
            noEndCap= self.error_format.No_endCap,
            val=err_Val,
            minus=MinusNumDataSource,
            plus=PlusNumeDataSource,
        )

    @check_method_input('')
    @staticmethod
    def add_trendline_to_serie(serie: opxl_XYSeries, TL: TrendLine):
        """
        serie: serie = series[i]
        """
        serie.trendline = opxl_Trendline(
            trendlineType=TL.trendline_Type,
            dispEq=TL.disp_Eq,
            dispRSqr=TL.disp_RSqr,
            intercept=TL.intercept,
            order=TL.order,
            period=TL.period,
        )

    @check_method_input('')
    @staticmethod
    def fuse2charts(
        ch_1: opxl_ScatterChart | opxl_BarChart | opxl_LineChart,
        ch_2: opxl_ScatterChart | opxl_BarChart | opxl_LineChart,
        worksheet:opxl_worksheet,
        position: str,
    ):
        """
        Display y-axis of the second chart on the right by setting it to cross the x-axis at its maximum
        """
        if not type(ch_1) == type(ch_2):
            raise TypeError("Chart type error")

        c = copy(ch_1)
        c2 = copy(ch_2)
        c.y_axis.crosses = "max"
        c += c2

        if not CellsCoordinates._isValidExcelWorksheet(worksheet):
            raise TypeError("Worksheet type error")

        if not CellsCoordinates._isValidExcelCell(position):
            raise Exception('Cell coordinate not valid')

        worksheet.add_chart(c, position)

        return c

    @staticmethod
    def _setSerieValue(axis_datas: DMSelect) -> opxl_chart_Reference:
        axis = CellsCoordinates.set_up_selection(axis_datas)

        values = opxl_chart_Reference(
            axis.worksheet,
            min_col=axis.columns_range.start,
            min_row=axis.rows_range.start,
            max_col=axis.columns_range.stop - 1,
            max_row=axis.rows_range.stop - 1,
        )

        return values
