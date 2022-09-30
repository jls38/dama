import logging
import os
#from Procedures.CSV_to_Excel_example  import configuration_1 as CSV_conf_1
#from Procedures.Data_to_Excel_example  import configuration_1 as Data_conf_1
#from Procedures.Nist_template import Generator as NIST_gen
from Procedures.Excel_format_example import workbook_format_1 as wb_f1
#from Procedures.graph import plot_graphe as plot

def main ():
    log_filename = os.path.join(os.getcwd(), 'logs', 'recette_1.log')
    logging.basicConfig(filename=log_filename,filemode= 'w', level=logging.INFO)
    log = logging.getLogger(__name__)

    log.info('Started')

    """Recette appliquée"""
    '''
    CSV_conf_1().recette()
    Data_conf_1().recette()
    
    plot()
    NIST_gen().recette()
    '''
    
    wb_f1().apply()


    log.info('Finiched')

if __name__ == "__main__":
    main()