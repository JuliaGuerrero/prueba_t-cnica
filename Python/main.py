# Autor: Julia Marinel Guerrero Obando
# Fecha: 2024/07/13
# Descripción: Script para combinar varios archivos Excel en uno solo, aplicando un tratamiento a los datos.
# Para: Proyecto de prueba para la empresa SkaloTek

#################### LEEME ####################
# Para poder ejecutar este script es necesario tener instalado Python y las librerías pandas, openpyxl, os.
# pip install pandas openpyxl os

import pandas as pd
import os

# Lista de archivos Excel
file_paths = ['BD/ENERO 2023.xlsx', 'BD/FEBRERO 2023.xlsx', 'BD/MARZO 2023.xlsx', 'BD/ABRIL 2023.xlsx', 'BD/MAYO 2023.xlsx', 'BD/JUNIO 2023.xlsx']

# Lista para almacenar los DataFrames
dataframes = []

# Función para convertir las fechas del 2024 al 2023
def convert_year_to_2023(date_str):
    if pd.isna(date_str):
        return date_str
    try:
        date = pd.to_datetime(date_str)
        if date.year == 2024:
            return date.replace(year=2023)
    except ValueError:
        return date_str
    return date_str

# Procesar cada archivo Excel
for file_path in file_paths:
    if os.path.exists(file_path):
        # Leer el archivo Excel
        df = pd.read_excel(file_path)

        # Mostrar el archivo que se está procesando
        print('Se está procesando el archivo: ', file_path)
        
        # Reemplazar los valores vacíos en las columnas especificadas
        df['CATEGORIA SOCIO'].fillna('Desconocido', inplace=True)
        df['ACCION'].fillna('Desconocido', inplace=True)
        df['COMENTARIO SOCIO'].fillna('Sin comentario', inplace=True)
        df['COMENTARIO CLUB'].fillna('Sin comentario', inplace=True)
        df['PROPINA'].fillna('Desconocido', inplace=True)
        df['FORMA DE PAGO'].fillna('Desconocido', inplace=True)
        df['COMENTARIO'].fillna('Sin comentario', inplace=True)
        df['Mesa'].fillna('Desconocido', inplace=True)
        df['PROVEEDOR'].fillna('Desconocido', inplace=True)
        df['RESTAURANTE'].fillna('Desconocido', inplace=True)

        # Eliminar las filas que contengan "pru" en la columna "COMENTARIO SOCIO" con la intención de eliminar comentarios de prueba
        df = df[~df['COMENTARIO SOCIO'].str.contains('pru', case=False, na=False)]

        # Eliminar las columnas que están sin registros
        df.drop(columns=['PROVEEDOR', 'ESTADO TRANSACCION', 'MEDIO DE PAGO'], inplace=True)
        
        # Convertir las fechas en la columna 'FECHA/HORA ENTREGA' del 2024 al 2023
        df['FECHA/HORA ENTREGA'] = df['FECHA/HORA ENTREGA'].apply(convert_year_to_2023)
        
        # Agregar el DataFrame a la lista
        dataframes.append(df)
    else:
        print(f'Archivo {file_path} no encontrado.')

# Verificar si hay DataFrames para concatenar
if dataframes:

    print('Combinando los DataFrames...')

    # Concatenar todos los DataFrames en uno solo
    combined_df = pd.concat(dataframes, ignore_index=True)
    
    # Guardar el DataFrame combinado a un nuevo archivo Excel
    output_file_path = 'Datos Depurados y Combinados.xlsx'
    combined_df.to_excel(output_file_path, index=False)
    
    print(f'Archivo combinado guardado en {output_file_path}')
else:
    print('No se encontraron archivos para procesar.')
