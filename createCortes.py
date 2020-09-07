#Code used to create CORTES with Python

#==============================================================================
#Libraries.
#==============================================================================
from os import path
from openpyxl import Workbook, load_workbook
#import win32com.client as w32lc
#import os
#import pandas as pd

#Paths
thisFilePath = path.abspath(__file__)
thisFolderPath = path.dirname(thisFilePath)

#==============================================================================
#Class definition.
#==============================================================================
class createCortes:
    def __init__(self, month, year, trimester):
        self.month = month
        self.year = year
        self.trimester = trimester

#=================================================================================
    def runCreateCortes(self): 
#=================================================================================
        fileName = str(self.year) + '-' + str(self.month) + '_Restricciones_'+str(self.year)+'_T'+str(self.trimester)+'.xlsx'
        #wb = load_workbook(thisFolderPath+'\\'+fileName)
        print(fileName)


#==============================================================================
#Main.
#==============================================================================
def main(month, year, trimester):
    
    Cortes=createCortes(month, year, trimester)
    Cortes.runCreateCortes()


#==============================================================================
#CODE STARTS HERE!!!!!
#==============================================================================
if __name__ == "__main__":
    #-----------------------------------------
    # Preliminars.
    month = 7
    year = 2020
    trimester = 3

    #-----------------------------------------
    # run main.
    main(month, year, trimester)