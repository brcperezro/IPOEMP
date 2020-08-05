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


#==============================================================================
#Main.
#==============================================================================
def main():
    month, year=askForMonth()                       #Get month and year form user
    trimester = (int(month)-1)//3 +1                #Calculate trimester
    createRestrictionsTable(month, year, trimester) #Create Restrictions table from IPOEMP table


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