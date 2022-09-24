import logging
from os import path as os_path
from dataclasses import dataclass
from typing import Sequence

from pandas import read_excel as pd_read_excel, DataFrame as pd_DataFrame, read_csv as pd_read_csv

from DamaLib.common.excel.workbook.xlfile import xlfile
from DamaLib.common.decorators.check import check_method_input

log = logging.getLogger(__name__)

@dataclass
class xlsx_properties:
    """
    sheets: Sheet to use to extract Dataframe. Pass None to extract all sheets from workbook.
    Header: Row (0-indexed) to use for the column labels of the parsed DataFrame. If a list of integers is passed those row positions will be combined into a MultiIndex. Use None if there is no header.
    idx_col: Column (0-indexed) to use as the row labels of the DataFrame. Pass None if there is no such column. If a list is passed, those columns will be combined into a MultiIndex. If a subset of data is selected with usecols, index_col is based on the subset.
    cols: 
        - If None, then parse all columns.
        - If str, then indicates comma separated list of Excel column letters and column ranges (e.g. “A:E” or “A,C,E:F”). Ranges are inclusive of both sides.
        - If list of int, then indicates list of column numbers to be parsed.
        - If list of string, then indicates list of column names to be parsed.
    rows: Number of rows to parse.
    """
    sheets:int|str|list|None = None
    Header:int|Sequence[int]|None= None
    idx_col:int|Sequence[int]|None = None
    cols:int|str|list|None = None
    rows:int|None = None
    

@dataclass
class CSV_properties:
    """
    sep must be one of (None, ',', '\t')
    decimal must be one of ('.', ',')
    """
    sep:str|None = '\t'
    decimal:str = '.'

    def __post_init__(self):
        if not self.sep in (None, ',', '\t'):
            raise ValueError("sep must be one of (None, ',', '\t')")

        if not self.decimal in ('.',','):
            raise ValueError("decimal must be one of ('.', ',')")

class LoadDataframe (object):

    def __init__(self) -> None:
        self._pathList = ''
        self._xl_prop = xlsx_properties
        self._csv_prop = CSV_properties
    
    @property
    def xl_properties(self) -> xlsx_properties:
        return self._xl_prop

    @check_method_input(('',))
    @xl_properties.setter
    def xl_properties(self, prop:xlsx_properties) -> None:
        self._xl_prop = prop

    @property
    def csv_properties(self) -> CSV_properties:
        return self._csv_prop

    @check_method_input(('',))
    @csv_properties.setter
    def csv_properties(self, prop:CSV_properties) -> None:
        self._csv_prop = prop

    @property
    def pathList(self) -> str:
        return self._pathList

    @check_method_input(('',))
    @pathList.setter
    def pathList(self, filespath:str|list) -> None:
        if type(filespath) is str:
            if not os_path.exists(filespath):
                raise ValueError("File doesn't exist")
        elif type(filespath) is list:
            for path in filespath:
                if not os_path.exists(path):
                    raise ValueError("File doesn't exist")
        else:
            raise TypeError('Path not valid')
        
        self._pathList = filespath     

    def load (self)-> dict|pd_DataFrame:
        """
        Extract pandas.Dataframe from file(s) and generate a dictionary with filename as key and pandas.dataframe as item
        """
        if self.pathList == '':
            raise ValueError('pathList not difined')

        if type(self.pathList) is str:
            loaded_files= self._get_df(self.pathList)
        else:
            loaded_files = {os_path.basename(path) : self._get_df(path) for path in self.pathList}

        return loaded_files

    def _get_df(self, path:str) -> dict:
        ext = os_path.splitext(path) [1]

        match ext:
            case '.xlsx':
                df = self._xlsx(path)
            case '.txt':
                df = self._csv(path)
            case '.csv':
                df = self._csv(path)
            case _:
                raise ValueError('Impossible to load file: extension not valid')
        
        log.info(f'{os_path.basename(path)} loaded')
        return df

    def _xlsx(self, filepath:str) -> pd_DataFrame:
        sheetsList = xlfile.get_xls_sheetsList(os_path.relpath(filepath))

        if not self._isXlsxPropDefined():
            prop = xlsx_properties()
        else:
            prop = self.xl_properties

        if prop.sheets == None:
            sList = sheetsList
        elif type(prop.sheets) == list and all(sheet in prop.sheets for sheet in sheetsList):
            sList = prop.sheets
        elif type(prop.sheets) == str and prop.sheets in sheetsList:
            sList = [prop.sheets]
        elif type(prop.sheets) == int and prop.sheets < len(sheetsList):
            sList = [prop.sheets]
        else:
            raise TypeError('Sheet type error')
            
        df_List = [
            pd_read_excel(
                filepath, 
                sheet_name= sheet,
                header= prop.Header,
                index_col= prop.idx_col,
                usecols= prop.cols,
                ) for sheet in sList
            ]

        return df_List

    def _csv (self, filepath:str) -> pd_DataFrame:
        path = os_path.relpath(filepath)

        if not self._isCsvPropDefined():
            prop = CSV_properties()
        else:
            prop = self.csv_properties

        df = pd_read_csv(
            path,
            sep= prop.sep,
            decimal= prop.decimal
            )

        return df

    def _isXlsxPropDefined(self):
        try:
            print(self.xl_properties.Header)
            return True
        except:
            return False

    def _isCsvPropDefined(self):
        try:
            print(self.csv_properties.decimal)
            return True
        except:
            return False