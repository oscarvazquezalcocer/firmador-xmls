from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from tkinter.filedialog import askdirectory
import os

# Ruta RElativa de la imagen
ruta_script = os.path.dirname(os.path.abspath(__file__))
ruta_imagen = os.path.join(ruta_script, "firmas.png")

ruta_carpeta = askdirectory(title="Selecciona una carpeta")

if ruta_carpeta:
    # Obten  la lista de archivos en la carpeta
    archivos = os.listdir(ruta_carpeta)


    # Recorre cad archivo de la cartpeta
    for archivo in archivos:
  
        # Verifica que el archivo se de Excel
        if archivo.endswith('.xlsx'):
          
            # Construye la ruta completa del archivo
            ruta_archivo = os.path.join(ruta_carpeta, archivo)

            # Carga el Archivo Existente
            workbook = load_workbook(ruta_archivo)

            #Selecciona la hoja
            sheet = workbook.active

            # Carga la imagen desde un archivo
            imagen_firma = Image(ruta_imagen)

            # Define la ubicación de la imagen en la hoja de cálculo de Excel
            imagen_firma.anchor = 'D57'

            # Agrega la imagen a la hoja de cálculo
            sheet.add_image(imagen_firma)

            # Guarda el archivo de Excel
            workbook.save(ruta_archivo)
