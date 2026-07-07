import os
import pandas as pd
from pathlib import Path

# --- 1. CONFIGURACIÓN DE RUTAS ---
BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_DATOS = BASE_DIR / "Flujo" / "Input" / "01 Proyeccion de Ventas - Maxx_v2.01 (1).xlsx"
SALIDA_DATOS = BASE_DIR / "Flujo" / "Output" / "Master.csv"

# --- 2. VARIABLES A CAMBIAR ---
COLUMNA_BUSQUEDA = "D"          
NOMBRE_COL_INICIO = "Categoria"
NOMBRE_COL_FIN = "Cuota Inicial"

# <--- ¡AQUÍ ESTÁ EL CAMBIO CLAVE! Ahora es una lista con las hojas que quieres juntar
HOJAS_EXCEL = ["Ventas Dptos", "Vnt Est - dpsto"] 

def transformar_excel_a_csv():
    try:
        lista_dataframes = [] # Aquí guardaremos los pedazos de datos de cada hoja
        
        for hoja in HOJAS_EXCEL:
            print(f"\n📄 Escaneando el archivo en la hoja '{hoja}'...")
            
            try:
                # Leer el Excel sin encabezados
                df_crudo = pd.read_excel(ENTRADA_DATOS, sheet_name=hoja, header=None)
            except ValueError:
                print(f"⚠️ ADVERTENCIA: No se encontró la hoja '{hoja}' o está mal escrita. Saltando...")
                continue
            
            # Convertir letra a índice numérico
            indice_columna = ord(COLUMNA_BUSQUEDA.upper()) - 65 
            
            # Escanear la columna
            columna_escaneada = df_crudo[indice_columna].astype(str).str.strip()
            filas_encontradas = df_crudo[columna_escaneada == NOMBRE_COL_INICIO].index
            
            if len(filas_encontradas) > 0:
                fila_encabezado = filas_encontradas[0] 
                print(f"🎯 ¡Encontrado! '{NOMBRE_COL_INICIO}' está en la fila: {fila_encabezado + 1}")
                
                # Extraemos todos los nombres de esa fila como una lista limpia
                nombres_fila = df_crudo.iloc[fila_encabezado].astype(str).str.strip().tolist()
                
                # --- LÓGICA: Buscar por POSICIÓN NUMÉRICA ---
                try:
                    # Buscamos en qué número de columna están los encabezados que queremos
                    idx_inicio = nombres_fila.index(NOMBRE_COL_INICIO)
                    idx_fin = nombres_fila.index(NOMBRE_COL_FIN)
                    
                    # Descartamos la "basura" de arriba y asignamos los nombres
                    df_datos = df_crudo.iloc[fila_encabezado + 1:].reset_index(drop=True)
                    df_datos.columns = nombres_fila
                    
                    # Cortamos usando la posición numérica (.iloc) en lugar del nombre
                    df_resultado = df_datos.iloc[:, idx_inicio : idx_fin + 1]
                    
                    # Guardamos este dataframe en nuestra lista para juntarlo al final
                    lista_dataframes.append(df_resultado)
                    print(f"✅ Datos extraídos correctamente de '{hoja}'")
                    
                except ValueError:
                    print(f"❌ ERROR: En la hoja '{hoja}', se encontró la fila, pero no '{NOMBRE_COL_FIN}'")
                    
            else:
                print(f"❌ ERROR: No encontró la palabra '{NOMBRE_COL_INICIO}' en la columna {COLUMNA_BUSQUEDA} de '{hoja}'.")

        # --- 3. CONSOLIDACIÓN Y GUARDADO ---
        if lista_dataframes: # Si la lista no está vacía (es decir, si encontró datos)
            print("\n🔄 Juntando todas las hojas extraídas...")
            
            # Juntamos todos los dataframes verticalmente (uno debajo del otro)
            df_maestro = pd.concat(lista_dataframes, ignore_index=True)
            
            # Guardar en CSV
            SALIDA_DATOS.parent.mkdir(parents=True, exist_ok=True)
            df_maestro.to_csv(SALIDA_DATOS, index=False, encoding='utf-8-sig')
            
            print(f"🎉 ¡Proceso completado con éxito! Archivo consolidado guardado en: {SALIDA_DATOS.name}")
        else:
            print("\n❌ No se extrajeron datos de ninguna hoja. El archivo CSV no fue creado.")

    except Exception as e:
        print(f"❌ Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    transformar_excel_a_csv()