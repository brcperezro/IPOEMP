#==============================================================================
#Libraries.
#==============================================================================
from os import path
from win32com.client import Dispatch

#==============================================================================
#Functions.
#==============================================================================

#-----------------------------------------
# Get month to study from user.
def askForMonth():
    while True:
        try:
            monthInput = input('Ingrese el mes y año a analizar de la forma \n <mm-yyyy>\n')    #Get month and year from user
            month, year = monthInput.split('-')                                                 #Split into two variables
            break
        except:
            print('\n¡¡Valor no válido!! Ingresa un mes en el formato indicado\n')
    
    return month, year

#Create Restrictions table from IPOEMP table
def createRestrictionsTable(month, year):
    wbIPOEMP = xl.Workbooks.Open(Filename = thisFolderPath + "\\"+IPOEMPfilename)   #Load IPOEMP excel file
    wbRestricciones = xl.Workbooks.Add()                                            #Create new excel file

    wsIPOEMP = wbIPOEMP.Worksheets('Restricciones').Copy(Before=wbRestricciones.Worksheets(1)) #Copy 'Restricciones' page to new file
    
    wbIPOEMP.Close(False)           #Close IPOEMP file without saving changes
    wbRestricciones.Close(True, thisFolderPath + "\\"+year+"-"+ month+"_Restricciones.xlsx") #Close new file saving changes


#==============================================================================
#Main.
#==============================================================================
def main():
    month, year=askForMonth()
    createRestrictionsTable(month, year)



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
    IPOEMPfilename = 'TablaRestricciones.xlsx'

    #-----------------------------------------
    # run main.
    main()