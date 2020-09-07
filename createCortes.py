#Código crear cortes en python
from openpyxl import Workbook, load_workbook
import win32com.client as w32lc
import os
import pandas as pd
#separa la cadena de texto obtenido del archivo de excel para obtener las palabras claves del corte, ejemplo: "Paez 220/110 kV / Paez - Juanchito 220 kV", las
#palabras claves obtenidas serían: "Paez 220/110 kV" y Paez - Juanchito 220 kV
def separar(texto,iCorte):
    texto=texto.replace(" ","")
    cortes=[]
    a=0
    b=0
    cortes.append([])
    for rLetra in texto:
        if iCorte==None or iCorte=="-" or iCorte==" ":
            iNumTags="null"
            return cortes,iNumTags
            break

        if rLetra == "/" or rLetra=="+":
            if rLetra=="/" and texto[a-1].isdigit() and texto[a+1].isdigit():
                cortes[b].append(rLetra)
                a+=1
                continue
            else:
                b=b+1
                cortes.append([])
                a+=1
                continue
        
        cortes[b].append(rLetra)
        a=a+1
    iNumTags=b+1
    return cortes, iNumTags

#encuentra la palabra el valor del corte separado en la lista de claves, en caso de no encontrarlo marca "NOT FOUND", pero cuenta el corte para corregirlo manual
def encontrar(clave,maxcolum):
    p='NOT FOUND'
    pqc='NOT FOUND'
    q='NOT FOUND'
    qqc='NOT FOUND' 
    clave=clave.replace(" ","").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("\n","").upper().replace("KV","").replace(".","").replace("-","").replace("–","")

    for iFila in range(2,maxcolum+1):
        valor=(str(sheet2.cell(row=iFila,column=1).value).replace(" ","").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("\n","").upper()).replace("KV","").replace(".","").replace("-","").replace("–","")
        if clave==valor:
            p=sheet2.cell(row=iFila,column=2).value
            pqc=sheet2.cell(row=iFila,column=3).value
            q=sheet2.cell(row=iFila,column=4).value
            qqc=sheet2.cell(row=iFila,column=5).value
            break
 
    return p,pqc,q,qqc

#concatena los valores para guardar en el .csv
def escribir_csv(sP,sPQC,sQ,sQC,iCorte,iValorCorte,sArea,dfDatos):
   dfDatos=dfDatos.append({'P':sP,'Pcalidad':sPQC,'Q':sQ,'Qcalidad':sQC,'SubArea':sArea,'Corte':iCorte,'Pmax':iValorCorte},ignore_index=True)
   return dfDatos



#%% main
#ruta=('d:\\mis documentos\Codigo Python\\')
ruta= os.path.dirname(os.path.realpath('__file__'))+"\\"
EXCEL='2020-08_Restricciones_2020_T3.xlsx'
libro=load_workbook(ruta+EXCEL)
diccionario=load_workbook(ruta+'diccionario.xlsx')
dicDatos=[]
dfDatos=pd.DataFrame([],columns=['P','Pcalidad','Q','Qcalidad','SubArea','Corte','Pmax'])
sheet=libro.active
sheet2=diccionario.active
columna=sheet.max_row
iCantidadCortes=1 #Contador de cortes para marcar en el .csv
for iFila in range(2,columna+1):
    texto=sheet.cell(row=iFila,column=3).value
    cortes,iNumTags=separar(texto,sheet.cell(row=iFila,column=7).value)
    if iNumTags=="null":
        continue
    for iColumna in range(0,iNumTags):
        clave=''.join(cortes[iColumna]) #concatenar
        sP,sPQC,sQ,sQC=encontrar(clave,sheet2.max_row)
        iValorCorte=sheet.cell(row=iFila,column=7).value
        sArea=sheet.cell(row=iFila,column=1).value
        dfDatos=escribir_csv(sP,sPQC,sQ,sQC,iCantidadCortes,iValorCorte,sArea,dfDatos)
    sheet.cell(row=iFila,column=8).value=str(iCantidadCortes)
    iCantidadCortes+=1
#guardar datos en el .csv
libro.save(ruta+EXCEL)
dfDatos.to_csv(ruta+'cortes_creados_2020-08'+'.csv',sep=';',index=False)




