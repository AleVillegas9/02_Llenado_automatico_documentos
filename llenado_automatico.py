# -*- coding: utf-8 -*-
"""


@author: Alejandro Villegas

"""

# Paso 1: Preparación de la base de datos

import pandas as pd

base = pd.read_csv(r"C:\Users\javal\OneDrive\Desktop\Portafolio 9\02_llenado_automático_documentos\base.csv")

# paso 2: Plantilla

plantilla = r"C:\Users\javal\OneDrive\Desktop\Portafolio 9\02_llenado_automático_documentos\examen.docx"

# Paso 3: Lista única de profesores

base['profesor'] = base['profesor'].astype(str)
profesores = list(base["profesor"].unique())


# Paso 4: Creación de funciones
from docx import Document
import os
from docx2pdf import convert
from pypdf import PdfWriter


#Funcion para rellenar el word.

def reemplazar_campos (doc, datos_fila):
    for p in doc.paragraphs:
        for key, value in datos_fila.items():
            marcador = f"{{{{{key}}}}}"
            if marcador in p.text:
                p.text = p.text.replace(marcador, str(value))
    return doc
            

#funcion para generar el examen individual

def generar_exmenes_individuales (plantilla_path, datos, carpeta_destino):
    df = datos
    os.makedirs(carpeta_destino, exist_ok= True)
    for _, fila in df.iterrows():
        doc_temp = Document(plantilla_path)
        doc_temp = reemplazar_campos(doc_temp, fila)
        nombre_archivo = f"{fila['lastname'].replace(' ','_')}.docx"
        ruta_salida = os.path.join(carpeta_destino, nombre_archivo)
        doc_temp.save(ruta_salida)
        print(f"🍀🙊🙊Examenes automaticos guardados con éxito en {ruta_salida}🙊🙊🍀")
        
    
#Funcion par unir pdf's. 

def unir_pdfs (carpeta, archivo_salida):
    merger = PdfWriter()
    archivos_pdf = [f for f in os.listdir(carpeta)  if 
f.lower().endswith(".pdf")]
    archivos_pdf.sort()
    for pdf in archivos_pdf:
        ruta_pdf = os.path.join(carpeta,pdf)
        merger.append(ruta_pdf)
        print(f"🎉 Añadido: {pdf}")
        merger.write(archivo_salida)
        merger.close()



# Paso 5: Iterar todo el proceso para cada profesor de la lista


for profesor in profesores:
    ruta_salida = fr"C:\Users\javal\OneDrive\Desktop\Portafolio 9\02_llenado_automático_documentos\examenes\{profesor}"
    base2 = base[base["profesor"]== f"{profesor}"] #Asegurate de crear una nueva base en este paso
    generar_exmenes_individuales(
        plantilla_path= plantilla,
        datos = base2, #y que se tome la nueva base como el argumento datos.
        carpeta_destino= ruta_salida)
    convert(ruta_salida,ruta_salida)
    ruta_pdf_completo = fr"C:\Users\javal\OneDrive\Desktop\Portafolio 9\02_llenado_automático_documentos\examenes\examenes_completos\{profesor}.pdf"
    unir_pdfs(ruta_salida, ruta_pdf_completo)



