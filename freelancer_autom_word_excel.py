''' 
====================================================================================
     Programa Python que automatiza un documento Word                                

     Job Description: frelancer.es

    Se trata sobre automatizar u docuemnto e word, sobre un reporte medioambiental, 
    cambiando palabras clave y llenando tablas con información que está recopilada 
    en exceles.

        Habilidades: Python, Excel, SQLite, Word
        Author: Eng Manuel Calistri
        mcalistri.freel@gmail.com
        Phone & Whatsapp: +058-426-903.5508
        ver 03, 21 April 2023.
===================================================================================
'''

from docxtpl import DocxTemplate    # para la plantilla
from datetime import datetime       # para la plantilla
import pandas as pd                 # ??? para la plantilla


import docx                             # new para informe final
from docx import Document               # new para informe final
doc_final= docx.Document()  # new para informe final.. creamos una instancia de Document

doc = DocxTemplate("frelancer_plantilla.docx")


Tiempo = [ 
    'Enero', 
    'Febrero', 
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre'
]


def get_IndexMes_TmpMax(t_lista, n_lista, Tiempo): 
        Tmax= t_lista[0]
        for i in range(n_lista):
            if (t_lista[i] > Tmax): # actualiza a max
                Tmax= t_lista[i]
                mes_Tmax= Tiempo[i]
        print("get_IndexMes_TmpMax(t_lista): la Temp max es: ", Tmax, "en el mes de:", mes_Tmax, "\n")
        return Tmax, mes_Tmax



def get_IndexMes_TmpMin(t_lista,  n_lista, Tiempo):
        
        Tmin= t_lista[0]
        print(t_lista)   # prueba........
        print("get_IndexMes_TmpMin(t_lista): ", t_lista, "\n")
        for i in range(n_lista):
            if (t_lista[i] <= Tmin): # actualiza a max
                Tmin= t_lista[i]
                mes_Tmin= Tiempo[i]
        print("get_IndexMes_TmpMin(t_lista): la Temp min es: ", Tmin, "en el mes de:", mes_Tmin, "\n")
        return Tmin, mes_Tmin



'''  =====  aqui cargamos lista de estaciones Excel desde archivo .csv ===='''
df_est= pd.read_csv('excel_lista_estaciones_ambientales.csv', encoding='latin-1')   #  df_est: estaciones ambientales
n=0
ultimo= False
for index_est, fila_est in df_est.iterrows():   # aqui extrae una fila (info) de cada/estacion ambiental
    ''' aqui  construimos el nombre del archivo con su tabla de datos'''
    print(fila_est)
    
    codigo_lista_fcsv= fila_est['codigo_lista']
    # print("\n codigo_lista_fcsv= ", codigo_lista_fcsv)

    ubicación_lista_fcsv= fila_est['ubicación_lista']
    # print("\n ubicación_lista_fcsv= ", ubicación_lista_fcsv)

    nom_data_csv= fila_est['nombre_csv_lista'] + ".csv"
    # print("\n nom_data_csv= ", nom_data_csv)
    # break

    n += 1  # la prueba se hace para 10 estaciones....
    
    if (n == 10):
        ultimo= True

    if (n >10 ):
        break
    
    '''  =====  aqui cargamos data Excel desde archivo .csv ===='''
    df= pd.read_csv(nom_data_csv, encoding='latin-1')   #  df: data frame

    vTmin= []
    vTmax= []
    vTmed= []

    context_add= {}

    n_lista= 0
    for index, fila in df.iterrows():   # aqui extrae una fila de la tabla

        vTmin.append([])
        vTmax.append([])
        vTmed.append([])

        # print(index)

        vTmax[index]= fila['Temp_max']
        print(vTmax[index])
        
        vTmin[index]= fila['Temp_min']
        print(vTmin[index])

        vTmed[index]= fila['Temp_media']
        print(vTmed[index], "\n")
        
        
        strTmax= 'T' + str( index ) + '_' + '1'
        print(strTmax, "= ", vTmax[index])

        strTmin= 'T' + str( index ) + '_' + '2'
        print(strTmin, "= ", vTmin[index])
        
        strTmed= 'T' + str( index ) + '_' + '3'
        print(strTmed, "= ", vTmed[index], "\n")
        
        
        context = {   # este es el 2do diccionario
            strTmax: str(vTmax[index]),       # aqui extrae cada columna de la fila
            strTmin: str(vTmin[index]),
            strTmed: str(vTmed[index])
        }

        n_lista += 1
        context_add.update(context)   

    #  Ok aqui procesamos los datos de los vectores para determinar los datos: Tmax, Tmin, y sus meses correspondientes del parrafo en el doc
    Tmax=0
    mes_Tmax="ninguno"
    Tmax, mes_Tmax= get_IndexMes_TmpMax(vTmax, n_lista, Tiempo)   # prueba de la rutina.....

    Tmin=0
    mes_Tmin="ninguno"
    
    Tmin, mes_Tmin= get_IndexMes_TmpMin(vTmin, n_lista, Tiempo)   # prueba de la rutina.....


    # procesar el contexo_analisis, asignar los valores determinados de tem max, min y meses 

    '''  =====  aqui creamos data (encabezado del informe) ===='''
    TmaxGen = Tmax
    mes_TmaxGen = mes_Tmax

    TminGen = Tmin
    mes_TminGen = mes_Tmin

    TmedAnual = vTmed[index]

    context_analisis = {  # este es el 1er diccionario
                'index' : str(n),  
                'codigo_lista' : codigo_lista_fcsv,  
                'ubicación_lista' : ubicación_lista_fcsv,
                'TmaxGen' : str(TmaxGen),
                'mes_TmaxGen' : mes_TmaxGen,
                'TminGen' : str(TminGen),
                'mes_TminGen' : mes_TminGen,
                'TmedAnual' : str(TmedAnual)
    }  



    context_analisis.update(context_add)   # a la plantilla del encabezado le agrega plantilla de la tabla excel
    # print(context_analisis)

    doc.render(context_analisis)

    
    doc.save(f"doc_generado.docx")  # aqui genera un archivo para los datos de estacion actual

    if (not ultimo): 
         doc.add_page_break() # agrega un salto de pagina
      
    for element in doc.element.body:  # actualiza "doc_final.docx" agregandole "doc_generado.docx" 
        doc_final.element.body.append(element)
        
    doc_final.save("doc_final.docx")  # guarda cambios


    




        

