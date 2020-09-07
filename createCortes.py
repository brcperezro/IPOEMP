#Code used to create CORTES with Python

#==============================================================================
#Libraries.
#==============================================================================
from os import path
from openpyxl import Workbook, load_workbook
from pandas import DataFrame
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
    def __init__(self, month, year, trimester, dictName):
        self.month = month
        self.year = year
        self.trimester = trimester
        self.dictName = dictName
    
        #function used to split the text into different tags
    def separateTags(self, text, MWlimit):
        text=text.replace(" ", "")  #removes all the spaces
        cortes = []
        iCorte=0
        cortes.append([])
        for Letter, iLetter in zip(text, range(len(text))):
            #condition used to remove sobrecargas and sobretensiones
            if MWlimit==None or MWlimit=="-" or MWlimit==" ":
                iNumTags ="null"
                return cortes, iNumTags
            
            # "/" and "+" are commonly used to separate Tags
            if Letter == "/" or Letter == "+":
                # Checks if the "/" is for transformators ratio (and not to separate Tags)
                if Letter == "/" and text[iLetter-1].isdigit() and text[iLetter+1].isdigit():
                    cortes[iCorte].append(Letter)
                    continue
                # If it is "+" or "/" to separate Tags
                else:
                    iCorte+=1
                    cortes.append([])
                    continue
            #if the letter is not "/" nor "+"
            cortes[iCorte].append(Letter)   #appends the letter to the string
        iNumTags = iCorte+1
        return cortes, iNumTags

        #function used to split the text into different tags
    def separateTags2(self, text, MWlimit):
        text=text.replace(" ", "")  #removes all the spaces
        cortes = []
        iCorte=0
        cortes.append('')
        for Letter, iLetter in zip(text, range(len(text))):
            #condition used to remove sobrecargas and sobretensiones
            if MWlimit==None or MWlimit=="-" or MWlimit==" ":
                iNumTags ="null"
                return cortes, iNumTags
            
            # "/" and "+" are commonly used to separate Tags
            if Letter == "/" or Letter == "+":
                # Checks if the "/" is for transformators ratio (and not to separate Tags)
                if Letter == "/" and text[iLetter-1].isdigit() and text[iLetter+1].isdigit():
                    cortes[iCorte]+=Letter  #appends the letter to the string
                    continue
                # If it is "+" or "/" to separate Tags
                else:
                    iCorte+=1
                    cortes.append('')   #Creates new string
                    continue
            #if the letter is not "/" nor "+"
            cortes[iCorte]+=Letter   #appends the letter to the string
        iNumTags = iCorte+1
        return cortes, iNumTags

#=================================================================================
    def runCreateCortes(self): 
#=================================================================================
        fileName = str(self.year) + '-' + str(self.month) + '_Restricciones_'+str(self.year)+'_T'+str(self.trimester)+'.xlsx'
        wbRestric = load_workbook(thisFolderPath+'\\'+fileName)     #Load Restricciones workbook
        wbDict = load_workbook(thisFolderPath+'\\'+fileName)        #Load Diccionario workbook
        dicData = []
        dfData = DataFrame([],columns=['P','Pcalidad','Q','Qcalidad','SubArea','Corte','Pmax'])
        wsRestric = wbRestric.active        #Get active worksheet on Restricciones workbook
        wsDict = wbDict.active              #Get active worksheet on Diccionario workbook
        column = wsRestric.max_row          #Get the max column value on Restricciones table
        iCantCortes = 1                     #Counter of Cortes to mark in the .csv file


        for iRow in range (2,3):#(2, column+1):
            text = wsRestric.cell(row=iRow, column=3).value
            Cortes, iNumTags = self.separateTags(text, wsRestric.cell(row=iRow, column=7).value) #gets original strings
            Cortes2, iNumTags2 = self.separateTags2(text, wsRestric.cell(row=iRow, column=7).value) #gets strings concatenated


#==============================================================================
#Main.
#==============================================================================
def main():
    
    Cortes=createCortes(month, year, trimester, dictName)       #Creates class named Cortes
    Cortes.runCreateCortes()    #Runs the function that creates cortes


#==============================================================================
#CODE STARTS HERE!!!!!
#==============================================================================
if __name__ == "__main__":
    #-----------------------------------------
    # Preliminars.
    month = '03'           #Month to analyse
    year = '2020'         #Year to analyse
    trimester = (int(month)-1)//3 +1        #Trimester of the year
    dictName = 'diccionario.xlsx'           #Dictionary filename

    #-----------------------------------------
    # run main.
    main()