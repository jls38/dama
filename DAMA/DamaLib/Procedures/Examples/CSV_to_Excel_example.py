import logging 
from os import path as os_path

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.filesIO.createfile import CreateFile, xlsx_properties as create_xl_prop
from DamaLib.common.filesIO.loadDataframe import LoadDataframe, CSV_properties as load_csv_prop
from DamaLib.common.excel.workbook.xlfile import xlfile

log = logging.getLogger(__name__)

class configuration_1 (object):
    def __init__(self) -> None:
        """Fichier(s) importé(s)"""
        self.filesnames = []
        self.df = None #Pandas Dataframes
        """Fichier(s) sortie(s)"""
        self.out_fpath = ""

    def recette (self):
        """Application de la recette"""
        log.info("Debut recette")

        self.f_1()

        log.info("Fin recette")

    def f_1 (self) -> None:
        log.info('f_1: Start')
        """Importation des feuilles de calcul"""
        #Get dictionary with filesname as key and df as item
        LoadDataframe.pathList = GetPathFromDialog().openfiles('Load files')
        filesDict = LoadDataframe().load()

        #Get files' names
        self.filesnames = [key for key in filesDict] 

        #Get first pd_DataFrame
        self.df = filesDict[self.filesnames[0]]
        print(self.df)

        """Fusion des données dans un Fichier Excel"""
        #Create Excel file as output
        CreateFile.filepath = GetPathFromDialog().saveas('Create file')
        self.out_fpath = CreateFile().create()
        #Add data on Excel file
        sheet = os_path.splitext(self.filesnames[0])[0] #Defined sheet name as file name without extension
        xlfile.add_to_xlsx(self.out_fpath, sheet, self.df,0,0)
        
        #Delete initial sheet generated while file created
        xlfile.remove_sheet(self.out_fpath, 'Sheet1')

        log.info('f_1 : Done')       