import logging
from copy import copy
from dataclasses import dataclass
from operator import xor

from openpyxl.worksheet.worksheet import Worksheet as opxl_worksheet
from openpyxl.chart.scatter_chart import ScatterChart as opxl_ScatterChart
from openpyxl.chart.line_chart import LineChart as opxl_LineChart
from openpyxl.chart.bar_chart import BarChart as opxl_BarChart
from openpyxl.chart.series import XYSeries as opxl_XYSeries
from openpyxl.chart.series_factory import SeriesFactory as opxl_chart_SeriesFactory
from openpyxl.chart.reference import Reference as opxl_chart_Reference
from openpyxl.chart.error_bar import ErrorBars as opxl_ErrorBars
from openpyxl.chart.data_source import NumDataSource as opxl_NumDataSource, NumRef as opxl_NumRef
from openpyxl.chart.trendline import Trendline as opxl_Trendline

from DamaLib.common.excel.sheetformat.Coordinates import cells_coordinates
from DamaLib.common.utils.color import color_converter as cmn_color_conv
from DamaLib.common.decorators.check import check_method_input
from DamaLib.common.decorators.apply_to_dec import for_all_methods
from DamaLib.common.decorators.debug import DebugClass


log = logging.getLogger(__name__)

@dataclass
class ChartFormat:
    title:str = 'Title'
    style:int = 13
    x_axis_title:str = 'x_axis title'
    y_axis_title:str = 'y_axis tile'
    x_axis_majorGridlines = None
    y_axis_majorGridlines = None

@dataclass
class SeriesData:
    """
    _axis: [_axis_start_cell:str, _axis_end_cell:str]
    X_axis: only one row/column can be selected
    """
    x_ws:opxl_worksheet
    X_axis:list
    y_ws:opxl_worksheet
    Y_axis:list
    title_from_data:bool=True

@dataclass
class SerieFormat:
    serie_title:str = None
    marker_shape:str = None
    marker_size:int|float = None
    marker_color:str = None
    hide_lines:bool = True

    def __post_init__ (self):
        if self.marker_color != None:
            self.marker_color = cmn_color_conv(self.marker_color).hexaCode() 

@dataclass
class ErrorData:
    """
    cust_Data = [start_cell:str, end_cell:str]
    """
    err_Val:int|float = None
    cust_Errmin_ws:opxl_worksheet = None
    cust_Errmin_Data:list = None
    cust_Errmax_ws:opxl_worksheet = None
    cust_Errmax_Data:list = None

    def __post_init__(self):
        if xor(self.cust_Errmax_Data is not None, self.cust_Errmax_ws is not None):
            raise ValueError('Value missing')

        if xor(self.cust_Errmin_Data is not None, self.cust_Errmin_ws is not None):
            raise ValueError('Value missing')

@dataclass
class ErrorFormat:
    """
    err_Dir: must be one of {'x', 'y'}
    err_Bar_Type: must be one of {'minus','both','plus'}
    err_val_type: must be one of {'fixedVal', 'cust', 'stdDev', 'percentage', 'stdErr'}
    """
    err_Dir:str
    err_Bar_Type:str
    err_val_type:str
    No_endCap:bool=False

    def __post_init__(self):
        if self.err_Dir not in {'x', 'y'}:
            raise ValueError("err_Dir must be one of {'x', 'y'}")

        if self.err_Bar_Type not in {'minus','both','plus'}:
            raise ValueError("err_Bar_Type must be one of {'minus','both','plus'}")

        if self.err_val_type not in {'fixedVal', 'cust', 'stdDev', 'percentage', 'stdErr'}:
            raise ValueError("err_val_type must be one of {'fixedVal', 'cust', 'stdDev', 'percentage', 'stdErr'}")

@dataclass
class TrendLine:
    """
    trenline_type: must be one of {'poly', 'power', 'exp', 'log', 'movingAvg', 'linear'}
    """
    trendline_Type:str
    disp_Eq:bool = False
    disp_RSqr:bool = False
    intercept:int|float|None = None
    order:int|None = None
    period:int|None = None

    def __post_init__(self):
        if self.trendline_Type not in {'poly', 'power', 'exp', 'log', 'movingAvg', 'linear'}:
            raise ValueError("trenline_type: must be one of {'poly', 'power', 'exp', 'log', 'movingAvg', 'linear'}")

#@DebugClass('__init__')
@for_all_methods(check_method_input(('',)), '')
class charts (cells_coordinates):
    def __init__(self, worksheet:opxl_worksheet) -> None:
        super().__init__(worksheet)
        
    def create_chart (self, chart_type:str, position:str) -> opxl_ScatterChart | opxl_BarChart | opxl_LineChart:
        """
        chart_type must be one of('LineChart', 'BarChart', 'ScatterChart')
        Position : cell coordinate of chart's top-left corner
        """
        match chart_type:
            case 'ScatterChart' :
                c = opxl_ScatterChart()
            case 'BarChart':
                c = opxl_BarChart()
            case 'LineChart':
                c = opxl_LineChart()
            case _:
                raise ValueError('Chart type incorrect or none implemented')

        if self._isValidExcelCell(position) == False:
            raise ValueError('Chart position invalid')

        self.worksheet.add_chart(c, position)
        return c

    def set_chart_format(self, c:opxl_ScatterChart|opxl_BarChart|opxl_LineChart, chartFormat:ChartFormat):
        c.title = chartFormat.title
        c.style = chartFormat.style
        c.x_axis.title = chartFormat.x_axis_title
        c.y_axis.title = chartFormat.y_axis_title
        c.x_axis.majorGridlines = chartFormat.x_axis_majorGridlines
        c.y_axis.majorGridlines = chartFormat.y_axis_majorGridlines

    def add_serie_to_chart(self, c:opxl_ScatterChart|opxl_BarChart|opxl_LineChart, SData:SeriesData) -> opxl_XYSeries:
                
        #Get x values
        xvalues = self._setSerieValue(SData.x_ws, SData.X_axis)
        
        #Get y values
        yvalues = self._setSerieValue(SData.y_ws, SData.Y_axis)

        series = opxl_chart_SeriesFactory(yvalues, xvalues, title_from_data=SData.title_from_data)
        c.series.append(series)

    def get_series_from_chart(self, c:opxl_ScatterChart|opxl_BarChart|opxl_LineChart) -> opxl_XYSeries:
        """
        Return series from chart
        """
        return c.series

    def set_serie_format(self, serie:opxl_XYSeries, Sformat:SerieFormat) -> None:
        """
        serie = chart.series[i]
        """
        serie.title = Sformat.serie_title
        serie.marker.symbol = Sformat.marker_shape
        serie.marker.size = Sformat.marker_size
        serie.marker.graphicalProperties.solidFill = Sformat.marker_color # Marker filling
        serie.marker.graphicalProperties.line.solidFill = Sformat.marker_color # Marker outline
        serie.graphicalProperties.line.noFill = Sformat.hide_lines # hide lines

    def add_serie_error_bar(self, serie:opxl_XYSeries,errData:ErrorData, errFormat:ErrorFormat):
        """
        serie: series[i]

        if {'stdErr'} used, no values needs to be specified
        if {'fixedVal', 'stdDev', 'percentage'} used, specified err_Val
        if {'cust'} used, specified cust_Errmin_List and/or cust_Errmax_List
        """
        #Set error type
        MinusNumDataSource = None
        PlusNumeDataSource = None
        err_Val = None
        
        if errFormat.err_val_type in {'fixedVal', 'stdDev', 'percentage'}:
            err_Val = errData.err_Val

        if errFormat.err_val_type == 'cust':
            match errFormat.err_Bar_Type:
                case 'minus':
                    Minus = self._setSerieValue(errData.cust_Errmin_ws, errData.cust_Errmin_Data)
                    MinusNumDataSource = opxl_NumDataSource(opxl_NumRef(Minus))
                case 'plus':
                    Plus = self._setSerieValue(errData.cust_Errmax_ws, errData.cust_Errmax_Data)
                    PlusNumeDataSource = opxl_NumDataSource(opxl_NumRef(Plus))
                case 'both':
                    Minus = self._setSerieValue(errData.cust_Errmin_ws, errData.cust_Errmin_Data)
                    MinusNumDataSource = opxl_NumDataSource(opxl_NumRef(Minus))

                    Plus = self._setSerieValue(errData.cust_Errmax_ws, errData.cust_Errmax_Data)
                    PlusNumeDataSource = opxl_NumDataSource(opxl_NumRef(Plus))

        #Add error to serie
        serie.errBars = opxl_ErrorBars(
            errDir= errFormat.err_Dir, 
            errBarType= errFormat.err_Bar_Type,
            errValType= errFormat.err_val_type,
            noEndCap= errFormat.No_endCap,
            val= err_Val,
            minus= MinusNumDataSource,
            plus= PlusNumeDataSource
            )

    def add_trendline_to_serie (self, serie:opxl_XYSeries, TL:TrendLine):
        """
        serie: serie = series[i]
        """
        serie.trendline = opxl_Trendline(
            trendlineType= TL.trendline_Type,
            dispEq= TL.disp_Eq,
            dispRSqr= TL.disp_RSqr,
            intercept= TL.intercept,
            order= TL.order,
            period= TL.period
            )

    def fuse2charts (self, ch_1:opxl_ScatterChart|opxl_BarChart|opxl_LineChart, ch_2:opxl_ScatterChart|opxl_BarChart|opxl_LineChart, position:str):
        """
        Display y-axis of the second chart on the right by setting it to cross the x-axis at its maximum
        """
        if type(ch_1) != type(ch_2):
            raise TypeError('Chart type error')

        c = copy(ch_1)
        c2 = copy(ch_2)
        c.y_axis.crosses = "max"
        c += c2
        self.worksheet.add_chart(c, position)
                
        return c

    def _setSerieValue (self, ws:opxl_worksheet, cells_range:list) -> opxl_chart_Reference:
        self.worksheet = ws
        self.set_cells(cells_range)

        values = opxl_chart_Reference(
            self.worksheet,
            min_col= self.cols_range.start, 
            min_row= self.rows_range.start,
            max_col= self.cols_range.stop - 1, 
            max_row= self.rows_range.stop - 1
            )
        
        return values

