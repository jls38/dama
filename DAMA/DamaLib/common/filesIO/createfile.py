import logging
from os import path as os_path, remove as os_remove
from dataclasses import dataclass

from tkinter.messagebox import askyesno as tk_askyesno
from pandas import DataFrame as pd_DataFrame, ExcelWriter as pd_ExcelWriter

from DamaLib.common.filesIO._checkpath import is_path_exists_or_creatable_portable
from DamaLib.common.decorators.apply_to_dec import for_all_methods
from DamaLib.common.decorators.check import check_method_input

log = logging.getLogger(__name__)

@dataclass
class xlsx_properties:
    df:pd_DataFrame = pd_DataFrame()
    sheet:str = 'Sheet1'

class CreateFile (object):
    def __init__(self) -> None:
        self._filepath = ''
        self._relpath = ''
        self._overwrite = bool
        self._xlsx_prop = xlsx_properties

    @property
    def filepath(self) -> str:
        return self._filepath

    @check_method_input(('',))
    @filepath.setter
    def filepath(self, filepath:str) -> None:
        if not is_path_exists_or_creatable_portable(filepath):
            raise ValueError('File path given is not valid')
        self._filepath = filepath

    @property
    def xlsx_prop(self):
        return self._xlsx_prop
    
    @check_method_input(('',))
    @xlsx_prop.setter
    def xlsx_prop(self, prop:xlsx_properties):
        self._xlsx_prop = prop

    @property
    def overwrite(self):
        return self._overwrite

    @check_method_input(('',))
    def overwrite(self, answer:bool):
        self._overwrite = answer

    def create(self) -> str:
        """Create a new file and return relative path"""
        if self.filepath == '':
            raise ValueError('Filepath not defined')

        #Get file extension
        ext = os_path.splitext(self.filepath) [1]

        if type(self.overwrite) != bool:
            self.overwrite = True

        if os_path.exists(self.filepath):
            if self.overwrite == True:
                answer = tk_askyesno('Overwriting file', f'Do you want to overwrite {os_path.basename(self._filepath)}? (No will end script)')
                if answer == False:
                    log.info('Script  end')
                    exit()
            os_remove(self.filepath)

        #Create file and return path
        match ext:
            case ".xlsx":
                path = self._xlsx()
            case ".txt":
                path = self._txt()
            case _:
                raise ValueError('Impossible to create file: extension not valid')

        return path
    
    def _xlsx(self) -> str:
        if not self._isXlsxPropDefined():
            prop = xlsx_properties()
        else:
            prop = self.xlsx_prop

        try:
            with pd_ExcelWriter(os_path.relpath(self.filepath)) as writer:
                prop.df.to_excel(
                    writer,
                    sheet_name= prop.sheet,
                    )
        except:
            raise Exception("File open, please close it to rewrite")
        

        return os_path.relpath(self.filepath)

    def _txt(self) -> str:
        f = open(self.filepath, 'w')
        f.close
        return os_path.relpath(self.filepath)

    def _isXlsxPropDefined(self):
        try:
            print(self.xlsx_prop.df)
            return True
        except:
            return False
