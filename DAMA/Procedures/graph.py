import matplotlib.pyplot as plt
from pandas import DataFrame
import numpy as np

from DamaLib.common.filesIO.getPathFromDialog import GetPathFromDialog, open_properties, saveas_properties
from DamaLib.common.filesIO.createfile import CreateFile, xlsx_properties as create_xl_prop
from DamaLib.common.filesIO.loadDataframe import LoadDataframe, xlsx_properties as load_xl_prop
from DamaLib.common.excel.workbook.xlfile import xlfile

def plot_graphe():
    GetPathFromDialog.open_prop = open_properties(default_extension='.xlsx')
    LoadDataframe.pathList = GetPathFromDialog().openfile('Load files')
    LoadDataframe.xl_properties = load_xl_prop(sheets='Feuille1',Header=0, cols=[0,1,3,4]) #Use if necessary to set properties
    df = LoadDataframe().load()[0]

    print('pandas df: ')
    print(df)
    df.to_numpy()
    print('numpy df: ')
    print(df['x'])


    plt.plot(df['x'], df['f(x)'], 'r--')
    plt.xlabel('x')
    plt.ylabel('f(x)')
    plt.show()


