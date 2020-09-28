#==============================================================================
#Libraries.
#==============================================================================
from os import path
from win32com.client import Dispatch
from createCortes import createCortes
from pandas import read_csv

#==============================================================================
#Functions.
#==============================================================================
#-----------------------------------------
# Get month to study from user.
def askForMonth():
    while True:
        try:
            monthInput = input('Ingresa el mes y año a analizar de la forma \n <mm-yyyy>\n')    #Get month and year from user
            month, year = monthInput.split('-')                                                 #Split into two variables
            break
        except:
            print('\n¡¡Valor no válido!! Ingresa un mes en el formato indicado\n')
    
    return month, year

#Create Restrictions table from IPOEMP table
def createRestrictionsTable(month, year, trimester):
    wbIPOEMP = xl.Workbooks.Open(Filename = thisFolderPath + "\\"+IPOEMPfilename)   #Load IPOEMP excel file
    wsIPOEMP = wbIPOEMP.Worksheets('Restricciones')                                 #Load 'Restricciones' sheet
    wbRestricciones = xl.Workbooks.Add()                                            #Create new excel file
    
    lastRow = wsIPOEMP.UsedRange.Rows.Count                                                                         #Get last row value    
    lastCol = [i for i in range(1,wsIPOEMP.UsedRange.Columns.Count) if wsIPOEMP.Cells(1,i).Value=="Corte\n[MW]"][0] #Get Col index for "Corte [MW]"

    wsIPOEMP.Range("A1", wsIPOEMP.Cells(lastRow, lastCol)).Copy()               #Copy the needed range
    wbRestricciones.ActiveSheet.Paste(wbRestricciones.ActiveSheet.Cells(1,1))   #Paste the needed range
    wbRestricciones.ActiveSheet.Cells(1, lastCol+1).Value = "Corte Python"      #Create new column
    wbRestricciones.ActiveSheet.Cells(1, lastCol+2).Value = "¿Superó corte?"    #Create new column
    wsIPOEMP.Cells(1, 1).Copy()     #Copy a cell to clear de clipboard (to avoid Excel message)
    wbIPOEMP.Close(False)           #Close IPOEMP file without saving changes
    wbRestricciones.Close(True, thisFolderPath + "\\"+year+"-"+ month+"_Restricciones_"+year+"_T"+str(trimester)+".xlsx") #Close new file saving changes

#Verifies if there are NOT FOUND tags. Returns True if it is ready; returns False to stop code
def verifyNotFoundTags(month, year):
    while True:
        # Load CSV file with tags
        try:
            pdTagsC = read_csv(thisFolderPath+'\\'+'Cortes_creados_'+year+'-'+month+'.csv', sep=',')
        except:
            input('Por favor cierra el archivo '+ 'Cortes_creados_'+year+'-'+month+'.csv' +' y presiona Enter para continuar')
            pdTagsC = read_csv(thisFolderPath+'\\'+'Cortes_creados_'+year+'-'+month+'.csv', sep=',')
        #Count NOT FOUND rows (Only checks on 'P' column)
        NotFoundcont = len(pdTagsC.loc[pdTagsC['P'] == 'NOT FOUND'])
        if NotFoundcont == 0:
            return True
        if NotFoundcont > 0: 
            ready =  input("No se encontraron " + str(NotFoundcont) + " cortes. Por favor corregirlos manualmente. \
            \nSi desea cancelar, presione 'N': ")
            #If user input is 'N', stop the code here
            if ready.upper() == 'N':
                return False


#==============================================================================
#Main.
#==============================================================================
def main():
    month, year=askForMonth()                       #Get month and year form user
    trimester = (int(month)-1)//3 +1                #Calculate trimester
    print('Creando tabla de restricciones...')
    createRestrictionsTable(month, year, trimester) #Create Restrictions table from IPOEMP table
    print('Creando Tags de los cortes...')
    cortes = createCortes(month, year, trimester, dictName) #Create Class taken from createCortes.py
    cortes.runCreateCortes() #Create file 'Cortes creados' from Dictionary
    bContinue = verifyNotFoundTags(month, year) #Verify if there are not found tags 
    

#==============================================================================
#CODE STARTS HERE!!!!!
#==============================================================================
if __name__ == "__main__":
    #-----------------------------------------
    # Preliminars.
    #Paths
    thisFilePath = path.abspath(__file__)
    thisFolderPath = path.dirname(thisFilePath)
    #Excel application
    xl = Dispatch("Excel.Application")
    #Needed files
    IPOEMPfilename = 'TablaRestricciones.xlsx'  #Restrictions filename
    dictName = 'diccionario.xlsx'               #Dictionary filename

    #-----------------------------------------
    # run main.
    main()