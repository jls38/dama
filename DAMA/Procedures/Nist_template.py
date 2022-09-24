import logging 
from os import path

from pandas import DataFrame as pd_DataFrame

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.filesIO.createfile import CreateFile, xlsx_properties as create_xl_prop
from DamaLib.common.filesIO.loadDataframe import CSV_properties, LoadDataframe, xlsx_properties as load_xl_prop
from DamaLib.common.excel.workbook.xlfile import xlfile

log = logging.getLogger(__name__)

class generator (object):
    def __init__(self) -> None:
        """Fichier(s) importé(s)"""
        self.df = None
        """Fichier(s) sortie(s)"""
        self.out_fpath = ""

    def recette (self):
        """Application de la recette"""
        log.info("Debut recette")

        self.generate_xl()

        log.info("Fin recette")

    def generate_xl (self) -> None:
        log.info('f_1: Start')
        """Importation des feuilles de calcul"""
        LoadDataframe.pathList = GetPathFromDialog().openfiles('Load files')
        LoadDataframe.csv_properties = CSV_properties(sep='\t', decimal='.')
        filesDict = LoadDataframe().load()

        """Fusion des données dans un nouveau fichier Excel"""
        CreateFile.filepath = GetPathFromDialog().saveas('Create file')
        self.out_fpath = CreateFile().create()

        for key in filesDict.keys():
            name = path.splitext(key)[0]
            print('name: ',name)
            xlfile.add_to_xlsx(self.out_fpath, name, filesDict[key], 0, 0, Header=True)

        xlfile.remove_sheet(self.out_fpath, 'Sheet1')

        log.info('f_1: Done')

    def f_2 (self) -> None:
        log.info('f_2: Start')
        


        log.info('f_2: Done')

        