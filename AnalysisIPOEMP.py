#Code used to run IPOEMP analysis

#==============================================================================
#Libraries.
#==============================================================================
from os import path, makedirs
from calendar import monthrange
from pandas import DataFrame
from numpy import genfromtxt

#Paths
thisFilePath = path.abspath(__file__)
thisFolderPath = path.dirname(thisFilePath)

#==============================================================================
#Class definition.
#==============================================================================
class obIPOEMP:
    def __init__(self, month, year, trimester, saveFilesPath, sampling, firstDay, lastDay):
        self.month = month
        self.year = year
        self.trimester = trimester
        self.saveFilesPath = saveFilesPath
        self.sampling = sampling
    
    def CreateFolder(self, directory):
        try:
            if not path.exists(directory):
                makedirs(directory)
        except OSError:
            print('Error creando directorio ' + directory)

#=================================================================================
    def analysisIPOEMP(self):
#=================================================================================
        self.CreateFolder(self.saveFilesPath)
        Punits = Pgenerators = Qgenerators = QcapAbsorb = QcapGenerate = []
        Tini = str(firstDay) + '/' + str(month) + '/' + str(year) + ' 00:00:00'
        Tfin = str(lastDay) + '/' + str(month) + '/' + str(year) + ' 23:59:00'

        dfResults = DataFrame()
        
        #from pandas import read_csv
        #Tags2 = read_csv('cortes_creados_' + self.year + '-' + self.month + '.csv', sep=',', encoding='latin-1')
        #Tags2 = Tags2.to_numpy()

        Tags = genfromtxt('cortes_creados_' + self.year + '-' + self.month + '.csv', delimiter=',', dtype=str, skip_header=1)
        print('Wiiii')





#==============================================================================
#Main.
#==============================================================================
def main():
    
    IPOEMP = obIPOEMP(month, year, trimester, saveFilesPath, sampling, firstDay, lastDay)       #Creates class named IPOEMP
    IPOEMP.analysisIPOEMP()    #Runs the function that runs the IPOEMP analysis


#==============================================================================
#CODE STARTS HERE!!!!!
#==============================================================================
if __name__ == "__main__":
    #-----------------------------------------
    # Preliminars.
    month = '05'           #Month to analyze
    year = '2020'         #Year to analyze
    trimester = (int(month)-1)//3 +1        #Trimester of the year
    firstDay = '1'
    #lastDay = '31'
    lastDay = monthrange(int(year), int(month))[1]

    dictName = 'diccionario.xlsx'           #Dictionary filename
    saveFilesPath = 'D:\\CortesIPOEMP\\'
    sampling = "1m" #muestreo

    #-----------------------------------------
    # run main.
    main()