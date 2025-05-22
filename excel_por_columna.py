# 
# Script para separar filas de un archivo excel por los valores encontrados en una columna
#
# 1. Crear ambiente virtual:
# python -m venv venv
# source venv/bin/activate
# 
# 2. Instalar dependencias:
# pip install pandas openpyxl
# 
# 3. Correr script:
# python ruta/a/archivo/excel_por_columna.py
# 

import pandas as pd
import os

# Cargar el archivo Excel original
archivo_entrada = '/Users/juanse/Documents/buk/batman/grupouma/2025-05-19 INFORMACIÓN ITEMS 2025-01-01.xlsx'
archivo_salida = '/Users/juanse/Documents/buk/batman/grupouma/separado_por_empresa.xlsx'

# Fila donde se encuentran los encabezados de columna (0-indexado)
# Si los encabezados están en la primera fila, usar 0
# Si están en la segunda fila, usar 1, etc.
fila_encabezados = 5  # Puedes cambiar este valor según tu archivo

try:
    # Leer los datos especificando la fila de encabezados
    df = pd.read_excel(archivo_entrada, header=fila_encabezados)
    
    # Mostrar las columnas disponibles
    print("Columnas disponibles en el archivo:")
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")
    
    # Preguntar al usuario qué columna usar para agrupar
    columna_indice = int(input("\nIngrese el número de la columna para agrupar: "))
    if 0 <= columna_indice < len(df.columns):
        columna_clave = df.columns[columna_indice]
        print(f"Usando la columna: '{columna_clave}' para agrupar")
    else:
        # Si el índice está fuera de rango, usar la primera columna
        columna_clave = df.columns[0]
        print(f"Índice fuera de rango. Usando la columna: '{columna_clave}' para agrupar")
    
    # Crear un archivo Excel nuevo con varias hojas
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        # Asegurarse de que hay al menos un grupo
        grupos = df.groupby(columna_clave)
        grupos_count = len(df[columna_clave].unique())
        
        if grupos_count == 0:
            # Si no hay grupos, escribir todo el DataFrame en una hoja
            df.to_excel(writer, sheet_name='Todos', index=False)
            print("No se encontraron grupos. Todos los datos se escribieron en una sola hoja.")
        else:
            # Escribir cada grupo en una hoja separada
            print(f"Creando {grupos_count} hojas...")
            for valor, grupo in grupos:
                # Asegurarse de que el nombre de la hoja sea válido
                if valor is None or valor == '':
                    nombre_hoja = 'Sin_nombre'
                else:
                    nombre_hoja = str(valor)[:31]  # Los nombres de hoja deben tener máximo 31 caracteres
                    # Reemplazar caracteres no válidos
                    nombre_hoja = nombre_hoja.replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')
                
                grupo.to_excel(writer, sheet_name=nombre_hoja, index=False)
                print(f"  - Hoja creada: {nombre_hoja} ({len(grupo)} filas)")
    
    print(f'\nSe creó el archivo: {archivo_salida}')

except FileNotFoundError:
    print(f"Error: No se encontró el archivo {archivo_entrada}")
    print(f"Verifique que la ruta sea correcta: {os.path.abspath(archivo_entrada)}")
except Exception as e:
    print(f"Error: {str(e)}")