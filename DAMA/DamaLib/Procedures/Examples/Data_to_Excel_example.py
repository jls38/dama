import logging 

from pandas import DataFrame as pd_DataFrame

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.filesIO.createfile import CreateFile, xlsx_properties as create_xl_prop
from DamaLib.common.filesIO.loadDataframe import LoadDataframe, xlsx_properties as load_xl_prop
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
        self.f_2()

        log.info("Fin recette")

    def f_1 (self) -> None:
        log.info('f_1 : Start')
        """Importation des feuilles de calcul"""
        LoadDataframe.pathList = GetPathFromDialog().openfiles('Load files')
        #LoadDataframe.xl_properties = load_xl_prop()
        filesDict = LoadDataframe().load()
        
        self.filesnames = [key for key in filesDict] #Get files name
        self.df = filesDict[self.filesnames[0]]

        """Fusion des données dans un Fichier Excel"""
        CreateFile.filepath = GetPathFromDialog().saveas('Create file')
        self.out_fpath = CreateFile().create()

        xlfile.add_to_xlsx(self.out_fpath,'Feuille1', self.df[0],0,0, Header=False)
        xlfile.add_to_xlsx(self.out_fpath,'Feuille1', self.df[1],0,3, Header=False)
        xlfile.add_to_xlsx(self.out_fpath,'Feuille2', self.df[0],0,0, Header=False)

        xlfile.remove_sheet(self.out_fpath, 'Sheet1')

        log.info('f_1 : Done')

    def f_2 (self) -> None:
        log.info('f_2 : Start')

        """Calculs feuille 1"""
        #Set df_op with 2 columns and its names
        col_name_1 = 'f1 + f2'
        col_name_2= '2*(f1 + f2)'
        df_op = pd_DataFrame(columns=[col_name_1,col_name_2])
        #Set values to colomns
        df_op[col_name_1] = self.df[0].iloc[1:,1] + self.df[1].iloc[1:,1]
        df_op[col_name_2] = (2*(self.df[0].iloc[1:,1] + self.df[1].iloc[1:,1]))

        #Add a column to df_op
        x_temp = [2*self.df[0].iloc[1:,1] + self.df[1].iloc[1:,1]]
        df_op.insert(2,'2f1 + f2', x_temp[0])
        """Ecriture dans le fichier"""
        xlfile.add_to_xlsx(self.out_fpath,'Feuille1', df_op,0,6)

        log.info('f_2 : Done')

        