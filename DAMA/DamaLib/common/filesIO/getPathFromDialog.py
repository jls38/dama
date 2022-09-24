from dataclasses import dataclass
import logging
import tkinter as tk
from tkinter import filedialog as tk_fd
from typing import Iterable

from DamaLib.common.decorators.check import check_method_input

log = logging.getLogger(__name__)

@dataclass
class open_properties:
    default_extension:str = '*.*'
    filetypes:Iterable[tuple[str,str|list[str]|tuple[str, ...]]]|None = (
        ('Excel file (.xlsx)', '.xlsx'),
        ('Text file (.txt)', '.txt'),
        ('csv file (.csv)', '.csv'),
        ('csv from text (.txt)', '.txt'),
        ('All files', '*,*'))

@dataclass
class saveas_properties:
    default_extension:str = '*,*'
    filetypes:Iterable[tuple[str,str|list[str]|tuple[str, ...]]]|None = (
        ('Excel file (.xlsx)', '.xlsx'),
        ('Text file (.txt)', '.txt'),
        ('All files', '*,*'))

class GetPathFromDialog (object):
    def __init__(self) -> None:
        self._open_prop = open_properties
        self._saveas_prop = saveas_properties

    @property
    def saveas_prop(self) -> saveas_properties:
        return self._saveas_prop
    
    @check_method_input(('',))
    @saveas_prop.setter
    def saveas_prop(self, prop:saveas_properties) -> None:
        self._saveas_prop = prop

    @property
    def open_prop(self) -> open_properties:
        return self._open_prop
    
    @check_method_input(('',))
    @open_prop.setter
    def open_prop(self, prop:open_properties) -> None:
        self._open_prop = prop
    
    @check_method_input(('',))
    def openfiles(self, label:str) -> list[str]:
        if not self._isOpenPropDefined():
            prop = open_properties()
        else:
            prop = self.open_prop
        
        root = tk.Tk()
        root.withdraw()
        pathList = tk_fd.askopenfilenames(
            title= label,
            defaultextension= prop.default_extension,
            filetypes= prop.filetypes,
            parent= root, 
            )

        #Exit if cancel
        if not pathList:
            log.info('Arret du script')
            exit()

        return pathList

    @check_method_input(('',))
    def openfile(self, label:str) -> str:
        if not self._isOpenPropDefined():
            prop = open_properties()
        else:
            prop = self.open_prop
        
        root = tk.Tk()
        root.withdraw()
        path = tk_fd.askopenfilename(
            title= label,
            defaultextension= prop.default_extension,
            filetypes= prop.filetypes,
            parent= root, 
            )

        #Exit if cancel
        if not path:
            log.info('Arret du script')
            exit()
            
        return path

    @check_method_input(('',))
    def saveas (self, label:str)-> str:
        if not self._isSaveasPropDefined():
            prop = saveas_properties()
        else:
            prop = self.open_prop
        
        path = tk_fd.asksaveasfilename(
            title= label,
            defaultextension= prop.default_extension,
            filetypes= prop.filetypes
        )

        #Exit if cancel
        if not path:
            log.info('Arret du script')
            exit()

        return path
    
    def _isOpenPropDefined(self):
        try:
            print(self.open_prop.default_extension)
            return True
        except:
            return False

    def _isSaveasPropDefined(self):
        try:
            print(self.saveas_prop.default_extension)
            return True
        except:
            return False