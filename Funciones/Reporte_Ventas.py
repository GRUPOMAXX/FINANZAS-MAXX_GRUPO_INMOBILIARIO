import os
import pandas as pd
import numpy as np
from pathlib import Path

# --- CONFIGURACIÓN DE RUTAS ---
BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_DATOS = BASE_DIR / "Flujo" / "Input" / "Reporte de Ventas"
SALIDA_DATOS = BASE_DIR / "Flujo" / "Output" / "Ingresos.csv"
 

DIC_UNIDADES = {
    "Dpto": "Departamento",
    "Dpto / Estac / Dep": "Departamento",
    "Estac + Deposito": "Estacionamiento + Deposito",
    "Estac. Simple": "Estacionamiento Simple",
    "Estac + Dep": "Estacionamiento + Deposito",
    "estac": "Estacionamiento Simple",
    "Estac": "Estacionamiento Simple"
}

MAPEO_COLUMNAS = {
    "UI": ["UI"],
    "Dpto": ["DPTO"],
    "Estac": ["ESTAC"],
    "Deposito": ["DEPOSITO"],
    "Datos del Cliente": ["DATOS DEL CLIENTE","CLIENTES"], 
    "Precio Venta": ["PRECIO TOTAL VENTA MN", "TOTAL PRECIO VENTA"],
    "Precio Departamento": ["PRECIO VENTA DPTO", "PRECIO DPTO"],
    "Precio Estacionamiento": ["PRECIO VENTA ESTAC", "PRECIO ESTAC"],
    "Precio Deposito": ["PRECIO VENTA DEPOSITO", "PRECIO DEPOSITO"],
    "Precio Estac + Dep": ["PRECIO ESTAC + DEP"],
    "Fecha Pago": ["FECHA PAGO"], 
    "Importe Abonado": ["IMPORTE ABONADO SOLES", "IMPORTE ABONADO"], 
    "Status": ["STATUS", "Status"]
}

archivos_excel = list(ENTRADA_DATOS.glob("*.xlsx"))
dfs_procesados = []

for archivo_path in archivos_excel:
    print(f"Procesando: {archivo_path.name}...")
    
    nombre_proyecto = archivo_path.stem.upper()
    for palabra in ["CONTROL DE VENTAS", "CONTROL VENTAS", "CONTROL VENTA", "-"]:
        nombre_proyecto = nombre_proyecto.replace(palabra, "")
    nombre_proyecto = nombre_proyecto.strip() 
    
    try:
        df_crudo = pd.read_excel(archivo_path, header=None)
        fila_ui = df_crudo[df_crudo[1].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.upper() == 'UI'].index
        
        if not fila_ui.empty:
            idx_header = fila_ui[0]
            df = pd.read_excel(archivo_path, header=idx_header)
            
            df.columns = (df.columns.astype(str)
                          .str.replace(r'[\r\n]+', ' ', regex=True) 
                          .str.replace(r'\s+', ' ', regex=True)
                          .str.strip()
                          .str.upper())
            
            df_limpio = pd.DataFrame()
            df_limpio['Proyecto'] = [nombre_proyecto] * len(df)

            for nombre_pbi, palabras_clave in MAPEO_COLUMNAS.items():
                col_encontrada = next((c for c in df.columns if any(pk in c for pk in palabras_clave)), None)
                if col_encontrada:
                    df_limpio[nombre_pbi] = df[col_encontrada]
                else:
                    df_limpio[nombre_pbi] = np.nan

            # --- LÓGICA DE COMPLETADO POR BLOQUE ---
            if 'Dpto' in df_limpio.columns:
                # 1. Columnas a completar
                cols_para_copiar = [
                    "UI", "Dpto", "Estac", "Deposito", "Datos del Cliente",
                    "Precio Venta", "Precio Departamento", "Precio Estacionamiento", 
                    "Precio Deposito", "Precio Estac + Dep"
                ]
                cols_presentes = [c for c in cols_para_copiar if c in df_limpio.columns]

                # 2. Rellenamos el Dpto primero para crear el "bloque continuo"
                df_limpio['Dpto'] = df_limpio['Dpto'].ffill()
                
                # 3. Agrupamos por el Dpto ya rellenado para copiar el resto (Precios, Cliente, UI)
                # Esto asegura que si el 104 tiene 3 filas de pago, las 3 tengan toda la info
                df_limpio[cols_presentes] = df_limpio.groupby(['Proyecto', 'Dpto'])[cols_presentes].ffill()

            # --- STATUS Y FILTRO ---
            if 'Status' in df_limpio.columns:
                df_limpio['Status'] = df_limpio['Status'].fillna('Saldo')

            if 'UI' in df_limpio.columns:
                df_limpio['UI'] = df_limpio['UI'].replace(DIC_UNIDADES)
            
            # Solo conservamos filas con Fecha Pago
            df_limpio = df_limpio.dropna(subset=['Fecha Pago'])
            dfs_procesados.append(df_limpio)
                
    except Exception as e:
        print(f"❌ Error en {archivo_path.name}: {e}")

if dfs_procesados:
    bd_final = pd.concat(dfs_procesados, ignore_index=True)
    bd_final.to_csv(SALIDA_DATOS, sep=';', index=False, encoding='utf-8-sig', quoting=1)
    print("-" * 30)
    print(f"✅ ¡Corregido! Ahora los bloques como el 104 están completos.")
else:
    print("❌ No se generaron datos.")