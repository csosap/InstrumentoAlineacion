#!/usr/bin/env python
# coding: utf-8

# In[1]:


# Importar las bibliotecas necesarias
import pandas as pd
from rich.console import Console

# Inicializar el objeto Console para usar rich
console = Console()

# Cargar los datos desde un archivo Excel con múltiples hojas
archivo = 'Base_Datos_Marcos_Estrategicos2.xlsx'
df_subsectores = pd.read_excel(archivo, sheet_name='CAD_PALABRAS')  # Hoja con subsectores
df_ods = pd.read_excel(archivo, sheet_name='ODS_PALABRA')  # Hoja con ODSS

# Asegurarse de que la columna 'PALABRAS CLAVE' contenga solo cadenas
df_subsectores['PALABRAS CLAVE'] = df_subsectores['PALABRAS CLAVE'].apply(lambda x: x if isinstance(x, str) else '')
df_ods['PALABRAS CLAVE'] = df_ods['PALABRAS CLAVE'].apply(lambda x: x if isinstance(x, str) else '')

# Función para asociar sectores y subcategorías
def asociar_sectores_y_subcategorias(descripcion):
    descripcion_lower = descripcion.lower()
    conteo_coincidencias = {}
    conteo_subcategorias = {}

    for index, row in df_subsectores.iterrows():
        if isinstance(row['PALABRAS CLAVE'], str):
            palabras_clave = row['PALABRAS CLAVE'].lower().split(', ')
            coincidencias = sum(1 for palabra in palabras_clave if palabra in descripcion_lower)

            if coincidencias > 0:
                sector = row['SECTOR CAD']  # Asegúrate de que este nombre sea correcto
                conteo_coincidencias[sector] = conteo_coincidencias.get(sector, 0) + coincidencias
                
                subcategoria = row['SUBSECTOR CRS']  # Asegúrate de que este nombre sea correcto
                conteo_subcategorias[subcategoria] = conteo_subcategorias.get(subcategoria, 0) + coincidencias

    sectores_ordenados = sorted(conteo_coincidencias.items(), key=lambda item: item[1], reverse=True)
    subcategorias_ordenadas = sorted(conteo_subcategorias.items(), key=lambda item: item[1], reverse=True)

    return sectores_ordenados[:3], subcategorias_ordenadas[:3]

# Función para asociar ODS
def asociar_ods(descripcion):
    descripcion_lower = descripcion.lower()
    conteo_coincidencias = {}
    
    for index, row in df_ods.iterrows():
        if isinstance(row['PALABRAS CLAVE'], str):
            palabras_clave = row['PALABRAS CLAVE'].lower().split(', ')
            coincidencias = sum(1 for palabra in palabras_clave if palabra in descripcion_lower)
            
            if coincidencias > 0:
                objetivo = row['ODS']  # Asegúrate de que este nombre sea correcto
                conteo_coincidencias[objetivo] = conteo_coincidencias.get(objetivo, 0) + coincidencias

    objetivos_ordenados = sorted(conteo_coincidencias.items(), key=lambda item: item[1], reverse=True)
    return objetivos_ordenados[:3]

# Menú interactivo
def menu():
    console.print("Seleccione la opción de asociación:", style="bold green")
    console.print("1. Asociar subsectores de CAD")
    console.print("2. Asociar objetivos ODS")
    
    opcion = input("Ingrese el número de opción (1 o 2): ")
    
    if opcion in ['1', '2']:
        descripcion_proyecto = input("Ingresa el nombre del proyecto o descripción corta: ")
        
        if opcion == '1':
            sectores_asociados, subcategorias_asociadas = asociar_sectores_y_subcategorias(descripcion_proyecto)
            if sectores_asociados:
                console.print("[bold blue]Sectores asociados con mayor coincidencia:[/bold blue]")
                for sector, coincidencias in sectores_asociados:
                    console.print(f"{sector}: {coincidencias} coincidencias")
            else:
                console.print("No se encontraron sectores asociados.")

            if subcategorias_asociadas:
                console.print("[bold blue]Subcategorías asociadas con mayor coincidencia:[/bold blue]")
                for subcategoria, coincidencias in subcategorias_asociadas:
                    console.print(f"{subcategoria}: {coincidencias} coincidencias")
            else:
                console.print("No se encontraron subcategorías asociadas.")
        
        elif opcion == '2':
            objetivos_asociados = asociar_ods(descripcion_proyecto)
            if objetivos_asociados:
                console.print("[bold blue]Objetivos ODS asociados con mayor coincidencia:[/bold blue]")
                for objetivo, coincidencias in objetivos_asociados:
                    console.print(f"{objetivo}: {coincidencias} coincidencias")
            else:
                console.print("No se encontraron objetivos ODS asociados.")
    else:
        console.print("Opción no válida.")

# Ejecutar el menú
if __name__ == "__main__":
    menu()


# In[ ]:




