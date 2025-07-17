# -*- coding: utf-8 -*-
"""
Created on Fri Apr 25 11:04:19 2025

@author: bcastaneda
"""

# pip install pandas openpyxl
import pandas as pd
import os

# Ruta del archivo de Excel
archivo_excel = r"C:/Users/bcastaneda/Documents/Codigos/SARLAFT.xlsx"

# Cargar el archivo de Excel
df = pd.read_excel(archivo_excel)

# Número de filas por fragmento
tamano_fragmento = 5000

# Dividir el DataFrame en fragmentos
fragmentos = [df[i:i + tamano_fragmento] for i in range(0, df.shape[0], tamano_fragmento)]

# Directorio donde se guardarán los archivos
directorio_salida = r"K:\Dirección de Incentivos y Subsidios\INCENTIVO-ISA\ISA\Documentos de Apoyo Pago ISA\1 Sarlaft"

# Cambiar al directorio de salida
os.chdir(directorio_salida)

# Nombres de las columnas que quieres separar (ajústalo si los nombres reales son diferentes)
columna_a = 'NIT_ASEGURADO'  # Reemplaza con el nombre real de la columna A
columna_b = 'NOMBRE_ASEGURADO'  # Reemplaza con el nombre real de la columna B

# Procesar y guardar cada fragmento por separado
for i, fragmento in enumerate(fragmentos):
    # Extraer columnas específicas
    fragmento_a = fragmento[[columna_a]]
    fragmento_b = fragmento[[columna_b]]
    
    # Nombres de archivo
    archivo_a = f"{columna_a}_{i + 1}.xlsx"
    archivo_b = f"{columna_b}_{i + 1}.xlsx"
    
    # Guardar cada fragmento en su propio archivo
    fragmento_a.to_excel(archivo_a, index=False)
    fragmento_b.to_excel(archivo_b, index=False)
    
    print(f"Fragmentos guardados: {archivo_a}, {archivo_b}")
