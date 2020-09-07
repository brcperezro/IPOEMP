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
    def separateTags(self, text, iCorte):
        text=text.replace(" ", "")  #removes all the spaces
        cortes = []
        a=b=0
        cortes.append([])
        for rLetter in text:
            #condition used to remove sobrecargas and sobretensiones
            if iCorte==None or iCorte=="-" or iCorte==" ":
                iNumTags ="null"
                return cortes, iNumTags
                break
            
            # "/" and "+" are commonly used to separate Tags
            if rLetter == "/" or rLetter == "+":
                # Checks if the "/" is for transformators ratio (and not to separate Tags)
                if rLetter == "/" and text[a-1].isdigit() and text[a+1].isdigit():
                    cortes[b].append(rLetter)
                    a+=1
                    continue
                # If it is "+" or "/" to separate Tags
                else:
                    b+=1
                    cortes.append([])
                    a+=1
                    continue
            cortes[b].append(rLetter)
            a+=1
        iNumTags = b+1
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
            Cortes, iNumTags = self.separateTags(text, wsRestric.cell(row=iRow, column=7).value)
            print (Cortes, iNumTags)


# #laksjflasfdlkajsfdlkjaslkdf
#         for iFila in range(2,columna+1):
#             texto=sheet.cell(row=iFila,column=3).value
#             cortes,iNumTags=separar(texto,sheet.cell(row=iFila,column=7).value)
#             if iNumTags=="null":
#                 continue
#             for iColumna in range(0,iNumTags):
#                 clave=''.join(cortes[iColumna]) #concatenar
#                 sP,sPQC,sQ,sQC=encontrar(clave,sheet2.max_row)
#                 iValorCorte=sheet.cell(row=iFila,column=7).value
#                 sArea=sheet.cell(row=iFila,column=1).value
#                 dfDatos=escribir_csv(sP,sPQC,sQ,sQC,iCantidadCortes,iValorCorte,sArea,dfDatos)
#             sheet.cell(row=iFila,column=8).value=str(iCantidadCortes)
#             iCantidadCortes+=1

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