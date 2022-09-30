import logging

from DamaLib.common.utils.color import Color
from DamaLib.common.excel.sheetformat.Charts import Charts as xl_charts, SerieFormat, SeriesData, ErrorData, ErrorFormat, TrendLine, ChartFormat, DMSelect

log = logging.getLogger(__name__)

class style_1(object):

    def __init__() -> None:
        pass

    def create_ScatterChart (ws, chart_position:str):
        xl_charts.chart_format = ChartFormat()
        chart = xl_charts().create_chart('ScatterChart', ws, chart_position)

        log.info('Scatter chart Done (%s)', chart_position )
        return chart
    
    def add_serie_to_ScatterChart (chart_ws, chart, x_ws, x_start:str, x_end:str, y_ws, y_start:str, y_end:str):
        X = DMSelect(x_ws, x_start, x_end)
        Y = DMSelect(y_ws, y_start, y_end)
        Sdata = SeriesData(X_axis=X, Y_axis=Y)
        xl_charts.add_series_to_chart(chart, Sdata)

        serie_1 = xl_charts.get_series_from_chart(chart)[0]
        
        xl_charts.serie_format = SerieFormat(marker_shape='triangle')
        xl_charts().apply_format_to_serie(serie_1)

        tl = TrendLine('linear', disp_Eq=True, disp_RSqr=True)
        xl_charts.add_trendline_to_serie(serie_1, tl)

    def add_error_to_ScatterChart(chart_ws, serie_ws ,serie):
        errMin = DMSelect(min_ws=serie_ws, min_start='A1', min_end='A')
        errMax = DMSelect(max_ws=serie_ws, max_start='A1', max_end='A')
        xl_charts.error_format = ErrorFormat('x','both','fixedVal')
        
        xl_charts().add_to_serie_error_bars(serie, ErrorData(2, errMin, errMax))
        log.info('Add series to chart Done')

    def fuse2charts(ws, c1, c2, position):
        xl_charts.fuse2charts(c1, c2, ws, position)