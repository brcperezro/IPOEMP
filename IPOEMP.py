#==============================================================================
#Libraries.
#==============================================================================
from os import path

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
    

#==============================================================================
#Main.
#==============================================================================
def main():
    month, year=askForMonth()



#==============================================================================
#CODE STARTS HERE!!!!!
#==============================================================================
if __name__ == "__main__":
    #-----------------------------------------
    # Preliminars.
    #Paths
    thisFilePath = path.abspath(__file__)
    thisFolderPath = path.dirname(thisFilePath)

    ruta_mapeo = r'\\archivosxm\AseguramientoOperacion\08.Seguimiento_Postoperativo\04.Seguimiento_AGC\01.Seguimiento_Diario'
    ruta_resultados = r'\\archivosxm\AseguramientoOperacion\08.Seguimiento_Postoperativo\04.Seguimiento_AGC\01.Seguimiento_Diario'
    ruta_imagenes = r'\\archivosxm\AseguramientoOperacion\08.Seguimiento_Postoperativo\04.Seguimiento_AGC\01.Seguimiento_Diario\Imagenes'
    ruta_plantilla_ReporteAGC = r'\\archivosxm\AseguramientoOperacion\08.Seguimiento_Postoperativo\04.Seguimiento_AGC\03.Macros y Plantillas'
    Umbral_Unidades=[50,50,50,50,50]
    Umbral_GR=[50,50,50]
    Umbral_SC=[50,50,70]
    intervalo=4
    umbral_error_int_real=35
    umbral_error_int_prog=50
    umbral_error=5

    #-----------------------------------------
    # run main.
    main()