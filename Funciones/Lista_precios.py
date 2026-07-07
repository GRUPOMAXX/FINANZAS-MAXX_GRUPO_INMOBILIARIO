import pandas as pd
from pathlib import Path
import re

BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_LISTA_PRECIOS = BASE_DIR / "Flujo" / "Input" / "Lista_de_precios"
SALIDA_LISTA = BASE_DIR / "Flujo" / "Output" / "Lista_precios.csv"

def extraer_nombre_proyecto(nombre_archivo):
    match = re.search(r"Stock\s*-\s*(.*?)\s*\d{2}\.\d{2}\.\d{2}", nombre_archivo, re.IGNORECASE)
    return match.group(1).strip() if match else "DESCONOCIDO"

def encontrar_inicio_tabla(df):
    """
    Busca la fila donde empieza la tabla real (donde aparece 'UNIDADES')
    """
    for i in range(len(df)):
        fila = df.iloc[i].astype(str).str.upper()
        if fila.str.contains("UNIDADES").any():
            return i
    return None

def reconstruir_columnas(df_raw, fila_header):
    """
    Une encabezados combinados (2 filas)
    """
    header1 = df_raw.iloc[fila_header].fillna("")
    header2 = df_raw.iloc[fila_header + 1].fillna("")

    columnas = []
    for h1, h2 in zip(header1, header2):
        col = f"{h1} {h2}".strip().upper()
        columnas.append(col)

    return columnas

def procesar_archivo(ruta):
    nombre_proyecto = extraer_nombre_proyecto(ruta.name)

    df_raw = pd.read_excel(ruta, header=None)

    fila_inicio = encontrar_inicio_tabla(df_raw)

    if fila_inicio is None:
        print(f"No se encontró tabla en {ruta.name}")
        return None

    columnas = reconstruir_columnas(df_raw, fila_inicio)

    df = df_raw.iloc[fila_inicio + 2:].copy()
    df.columns = columnas

    # 🔥 FILTRAR SOLO LAS COLUMNAS QUE NECESITAS
    columnas_deseadas = {
        "UNIDADES (DPTOS)": None,
        "TECHADA": None,
        "LIBRE": None,
        "DORM": None,
        "PRECIO EN SOLES": None,
        "UBICACIÓN": None,
        "STATUS": None
    }

    for col in df.columns:
        col_upper = col.upper()
        if "UNIDADES" in col_upper:
            columnas_deseadas["UNIDADES (DPTOS)"] = col
        elif "TECHAD" in col_upper:
            columnas_deseadas["TECHADA"] = col
        elif "LIBRE" in col_upper:
            columnas_deseadas["LIBRE"] = col
        elif "DORM" in col_upper:
            columnas_deseadas["DORM"] = col
        elif "PRECIO" in col_upper and "SOLES" in col_upper:
            columnas_deseadas["PRECIO EN SOLES"] = col
        elif "UBIC" in col_upper:
            columnas_deseadas["UBICACIÓN"] = col
        elif "STATUS" in col_upper:
            columnas_deseadas["STATUS"] = col

    columnas_finales = {k: v for k, v in columnas_deseadas.items() if v is not None}

    df_final = df[list(columnas_finales.values())].copy()
    df_final.columns = list(columnas_finales.keys())

    df_final["PROYECTO"] = nombre_proyecto

    # eliminar filas vacías
    df_final = df_final.dropna(how="all")

    return df_final

def main():
    archivos = list(ENTRADA_LISTA_PRECIOS.glob("*.xlsx"))

    dataframes = []

    for archivo in archivos:
        print(f"Procesando: {archivo.name}")
        df = procesar_archivo(archivo)

        if df is not None:
            dataframes.append(df)

    if dataframes:
        df_total = pd.concat(dataframes, ignore_index=True)
        df_total.to_csv(SALIDA_LISTA, index=False, encoding="utf-8-sig")
        print("✅ Archivo generado correctamente")
    else:
        print("❌ No se generó data")

if __name__ == "__main__":
    main()