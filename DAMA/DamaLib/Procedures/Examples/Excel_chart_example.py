import logging

from DamaLib.common.utils.color import Color
from DamaLib.common.excel.sheetformat.Charts import charts as xl_charts, SerieFormat, SeriesData, ErrorData, ErrorFormat, TrendLine, ChartFormat

log = logging.getLogger(__name__)

class style_1(object):

    def __init__() -> None:
        pass

    def create_ScatterChart (ws, chart_position:str):
        chart = xl_charts(ws).create_chart('ScatterChart', chart_position)
        xl_charts(ws).set_chart_format(chart, ChartFormat())

        log.info('Scatter chart Done (%s)', chart_position )
        return chart
    
    def add_serie_to_ScatterChart (chart_ws, chart, x_ws, x_start:str, x_end:str, y_ws, y_start:str, y_end:str):
        X = [x_start, x_end]
        Y = [y_start, y_end]
        
        xl_charts(chart_ws).add_serie_to_chart(chart, SeriesData(x_ws, X, y_ws, Y))
        series = xl_charts(chart_ws).get_series_from_chart(chart)
        serie = series[0]
        
        Sformat_1 = SerieFormat(marker_shape='triangle')
        xl_charts(chart_ws).set_serie_format(serie, Sformat_1)

        tl = TrendLine('linear', True, True)
        xl_charts(chart_ws).add_trendline_to_serie(serie, tl)

    def add_error_to_ScatterChart(chart_ws, serie_ws ,serie):
        ErrorData(2, serie_ws, ['A1', 'A'], serie_ws, ['A1', 'A'])
        ErrorFormat('x','both','fixedVal')

        xl_charts(chart_ws).add_serie_error_bar(serie,'x','both','fixedVal', err_Val=2)
        log.info('Add series to chart Done')

    def fuse2charts(ws, c1, c2, position):
        xl_charts(ws).fuse2charts(c1, c2, position)