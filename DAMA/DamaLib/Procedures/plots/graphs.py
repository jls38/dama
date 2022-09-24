import matplotlib.pyplot as plt

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.filesIO.createfile import CreateFile, xlsx_properties as create_xl_prop
from DamaLib.common.filesIO.loadDataframe import LoadDataframe, xlsx_properties as load_xl_prop
from DamaLib.common.excel.workbook.xlfile import xlfile

def plot_graphe():
    LoadDataframe.pathList = GetPathFromDialog.openfile('Load files')
    load_xl_prop(sheet='Feuille1', cols=[1,2]) #Use if necessary to set properties
    df = LoadDataframe().load()[0]
    print(df)

if __name__ == "__main__":
    plot_graphe()