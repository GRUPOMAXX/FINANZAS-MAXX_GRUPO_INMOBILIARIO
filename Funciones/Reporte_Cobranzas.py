# Programa para Generar Reporte de Cobranzas
# Creado por Eduardo Miguel Huamani Acosta                                  06/07/26

import pandas as pd
from pathlib import Path
import re
import unicodedata

import Descarga_Cobranzas as descarga

FECHA_INICIO_REPORTE = "01/06/2024" # Fecha de Inicio del Reporte
FECHA_FIN_REPORTE = "30/06/2030"  # Fecha final

EJECUTAR_DESCARGA = True

BASE_DIR = Path(__file__).resolve().parent.parent

ENTRADA_LISTA_PRECIOS = BASE_DIR / "Flujo" / "Input" / "Cobranzas"
SALIDA_LISTA = BASE_DIR / "Flujo" / "Output"

ENTRADA_LISTA_PRECIOS.mkdir(parents=True, exist_ok=True)
SALIDA_LISTA.mkdir(parents=True, exist_ok=True)


PROYECTO = {
    "RESIDENCIAL PRADA": "Prada",
    "RESIDENCIAL BEYOND": "Beyond",
    "RESIDENCIAL VENECIA": "Venecia"
}


COLUMNAS_FECHA = [
    "Fecha_Programada",
    "FechaPago"
]


COLUMNAS_NUMERO_FIJAS = [
    "Monto_Cuota",
    "Monto_Cuota_Pagado",
    "SaldoPorPagarCuota"
]


COLUMNAS_NUMERO_CON_INDICE = [
    "PrecioLista",
    "PrecioVenta"
]


def ejecutar_descarga_cobranzas():
    descarga.FECHA_INICIO_REPORTE = FECHA_INICIO_REPORTE
    descarga.FECHA_FIN_REPORTE = FECHA_FIN_REPORTE

    descarga.BASE_DIR = BASE_DIR
    descarga.ENTRADA_COBRANZAS = ENTRADA_LISTA_PRECIOS
    descarga.ENTRADA_COBRANZAS.mkdir(parents=True, exist_ok=True)

    descarga.main()


def normalizar_texto(texto):
    texto = str(texto).strip().upper()

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return texto


def obtener_nombre_proyecto(proyecto):
    proyecto_normalizado = normalizar_texto(proyecto)

    for nombre_original, nombre_corto in PROYECTO.items():
        if normalizar_texto(nombre_original) == proyecto_normalizado:
            return nombre_corto

    return str(proyecto).strip()


def obtener_anio_archivo(nombre_archivo):
    coincidencia = re.search(r"(20\d{2})", nombre_archivo)

    if coincidencia:
        return coincidencia.group(1)

    return ""


def leer_csv_seguro(ruta_archivo):
    codificaciones = ["utf-8-sig", "utf-8", "latin1", "cp1252"]

    for encoding in codificaciones:
        try:
            df = pd.read_csv(
                ruta_archivo,
                sep=",",
                dtype=str,
                encoding=encoding,
                keep_default_na=False
            )

            return df

        except Exception:
            continue

    raise ValueError(f"No se pudo leer el archivo: {ruta_archivo}")


def obtener_columnas_union(archivos_csv):
    columnas_union = []
    columnas_vistas = set()

    for archivo in archivos_csv:
        df = leer_csv_seguro(archivo)

        columnas_archivo = list(df.columns)

        for columna in columnas_archivo:
            if columna not in columnas_vistas:
                columnas_union.append(columna)
                columnas_vistas.add(columna)

    return columnas_union


def ordenar_columnas_cobranzas(columnas_union):
    campos_inmueble = [
        "Tipo_Construccion",
        "Nombre_Tipo_Construccion",
        "TipoInmueble",
        "NroInmueble",
        "PrecioLista",
        "PrecioVenta"
    ]

    columnas_auxiliares = [
        "Archivo_Origen",
        "Anio_Descarga",
        "Fecha_Inicio_Reporte",
        "Fecha_Fin_Reporte"
    ]

    indices = set()
    columnas_a_reordenar = set()

    for columna in columnas_union:
        for campo in campos_inmueble:
            patron = rf"^{campo}_(\d+)$"
            coincidencia = re.match(patron, columna)

            if coincidencia:
                indice = int(coincidencia.group(1))
                indices.add(indice)
                columnas_a_reordenar.add(columna)

    columnas_inmueble_ordenadas = []

    for indice in sorted(indices):
        for campo in campos_inmueble:
            columna = f"{campo}_{indice}"

            if columna in columnas_union:
                columnas_inmueble_ordenadas.append(columna)

    if "Nombres_Titular" in columnas_union:
        columnas_a_reordenar.add("Nombres_Titular")
        columnas_inmueble_ordenadas.append("Nombres_Titular")

    posiciones = [
        i
        for i, columna in enumerate(columnas_union)
        if columna in columnas_a_reordenar
    ]

    if posiciones:
        posicion_insertar = min(posiciones)
    else:
        posicion_insertar = len(columnas_union)

    columnas_sin_bloque = [
        columna
        for columna in columnas_union
        if columna not in columnas_a_reordenar
    ]

    columnas_reordenadas = (
        columnas_sin_bloque[:posicion_insertar]
        + columnas_inmueble_ordenadas
        + columnas_sin_bloque[posicion_insertar:]
    )

    columnas_finales = [
        columna
        for columna in columnas_reordenadas
        if columna not in columnas_auxiliares
    ]

    columnas_finales = columnas_finales + columnas_auxiliares

    return columnas_finales


def convertir_columna_numero(serie):
    serie = serie.astype(str).str.strip()

    serie = serie.str.replace("S/", "", regex=False)
    serie = serie.str.replace("US$", "", regex=False)
    serie = serie.str.replace("$", "", regex=False)
    serie = serie.str.replace(" ", "", regex=False)

    serie = serie.str.replace(",", "", regex=False)

    serie = serie.replace("", pd.NA)

    return pd.to_numeric(serie, errors="coerce")


def es_columna_numero(columna):
    if columna in COLUMNAS_NUMERO_FIJAS:
        return True

    for campo in COLUMNAS_NUMERO_CON_INDICE:
        patron = rf"^{campo}_(\d+)$"

        if re.match(patron, columna):
            return True

    return False


def convertir_tipos_base(base):
    for columna in COLUMNAS_FECHA:
        if columna in base.columns:
            base[columna] = pd.to_datetime(
                base[columna],
                errors="coerce",
                dayfirst=True
            )

    for columna in base.columns:
        if es_columna_numero(columna):
            base[columna] = convertir_columna_numero(base[columna])

    return base


def clasificar_tipo_inmueble(tipo_inmueble):
    tipo = normalizar_texto(tipo_inmueble)

    if "DEPART" in tipo:
        return "Departamentos"

    if "ESTACION" in tipo:
        return "Estacionamientos"

    if "DEPOS" in tipo:
        return "Depositos"

    if "LOCAL" in tipo or "COMERCIAL" in tipo:
        return "Local Comercial"

    return ""


def generar_reporte_ventas_por_proyecto(base):
    if "Proyecto" not in base.columns:
        raise ValueError("No existe la columna Proyecto en la base consolidada.")

    columnas_tipo = [
        columna
        for columna in base.columns
        if re.match(r"^TipoInmueble_(\d+)$", columna)
    ]

    indices = []

    for columna in columnas_tipo:
        coincidencia = re.match(r"^TipoInmueble_(\d+)$", columna)

        if coincidencia:
            indices.append(int(coincidencia.group(1)))

    indices = sorted(set(indices))

    registros = []

    for indice in indices:
        columna_tipo = f"TipoInmueble_{indice}"
        columna_nro = f"NroInmueble_{indice}"

        if columna_tipo not in base.columns:
            continue

        if columna_nro not in base.columns:
            continue

        temporal = base[["Proyecto", columna_tipo, columna_nro]].copy()

        temporal.columns = [
            "Proyecto",
            "TipoInmueble",
            "NroInmueble"
        ]

        registros.append(temporal)

    columnas_reporte = [
        "Proyecto",
        "Departamentos",
        "Estacionamientos",
        "Depositos",
        "Local Comercial"
    ]

    if not registros:
        return pd.DataFrame(columns=columnas_reporte)

    base_larga = pd.concat(registros, ignore_index=True)

    base_larga["Proyecto"] = base_larga["Proyecto"].astype(str).str.strip()
    base_larga["Proyecto"] = base_larga["Proyecto"].apply(obtener_nombre_proyecto)

    base_larga["TipoInmueble"] = base_larga["TipoInmueble"].astype(str).str.strip()
    base_larga["NroInmueble"] = base_larga["NroInmueble"].astype(str).str.strip()

    base_larga = base_larga[
        (base_larga["Proyecto"] != "")
        & (base_larga["TipoInmueble"] != "")
        & (base_larga["NroInmueble"] != "")
    ]

    base_larga["Categoria"] = base_larga["TipoInmueble"].apply(clasificar_tipo_inmueble)

    base_larga = base_larga[
        base_larga["Categoria"] != ""
    ]

    base_larga = base_larga.drop_duplicates(
        subset=[
            "Proyecto",
            "Categoria",
            "NroInmueble"
        ]
    )

    reporte = (
        base_larga
        .groupby(["Proyecto", "Categoria"])["NroInmueble"]
        .nunique()
        .unstack(fill_value=0)
        .reset_index()
    )

    for columna in columnas_reporte:
        if columna not in reporte.columns:
            reporte[columna] = 0

    reporte = reporte[columnas_reporte]

    orden_proyectos = list(PROYECTO.values())

    otros_proyectos = [
        proyecto
        for proyecto in reporte["Proyecto"].tolist()
        if proyecto not in orden_proyectos
    ]

    orden_final = orden_proyectos + otros_proyectos

    reporte["Proyecto"] = pd.Categorical(
        reporte["Proyecto"],
        categories=orden_final,
        ordered=True
    )

    reporte = reporte.sort_values("Proyecto")
    reporte["Proyecto"] = reporte["Proyecto"].astype(str)

    return reporte


def aplicar_formatos_excel(ruta_excel, nombre_hoja):
    from openpyxl import load_workbook

    wb = load_workbook(ruta_excel)
    ws = wb[nombre_hoja]

    encabezados = {}

    for celda in ws[1]:
        encabezados[celda.value] = celda.column

    for columna in COLUMNAS_FECHA:
        if columna in encabezados:
            col_idx = encabezados[columna]

            for fila in range(2, ws.max_row + 1):
                celda = ws.cell(row=fila, column=col_idx)
                celda.number_format = "dd/mm/yyyy"

    for columna in encabezados:
        if es_columna_numero(columna):
            col_idx = encabezados[columna]

            for fila in range(2, ws.max_row + 1):
                celda = ws.cell(row=fila, column=col_idx)
                celda.number_format = "#,##0.00"

    ws.freeze_panes = "A2"

    wb.save(ruta_excel)


def aplicar_formatos_reporte_ventas(ruta_excel, nombre_hoja):
    from openpyxl import load_workbook

    wb = load_workbook(ruta_excel)
    ws = wb[nombre_hoja]

    for row in ws.iter_rows(min_row=2):
        for cell in row[1:]:
            cell.number_format = "0"

    ws.freeze_panes = "A2"

    wb.save(ruta_excel)


def consolidar_cobranzas():
    archivos_csv = sorted(ENTRADA_LISTA_PRECIOS.glob("Cobranzas_*.csv"))

    if not archivos_csv:
        raise FileNotFoundError(
            f"❌ No se encontraron archivos Cobranzas_*.csv en: {ENTRADA_LISTA_PRECIOS}"
        )

    print("✅ Archivos encontrados:")

    for archivo in archivos_csv:
        print("-", archivo.name)

    columnas_union = obtener_columnas_union(archivos_csv)

    columnas_finales = ordenar_columnas_cobranzas(columnas_union)

    bases = []

    for archivo in archivos_csv:
        df = leer_csv_seguro(archivo)

        anio_archivo = obtener_anio_archivo(archivo.name)

        df["Archivo_Origen"] = archivo.name
        df["Anio_Descarga"] = anio_archivo
        df["Fecha_Inicio_Reporte"] = FECHA_INICIO_REPORTE
        df["Fecha_Fin_Reporte"] = FECHA_FIN_REPORTE

        for columna in columnas_finales:
            if columna not in df.columns:
                df[columna] = ""

        df = df[columnas_finales]

        bases.append(df)

        print(f"{archivo.name}: {len(df)} filas, {len(df.columns)} columnas")

    base_consolidada = pd.concat(bases, ignore_index=True)

    for columna in columnas_finales:
        if columna not in base_consolidada.columns:
            base_consolidada[columna] = ""

    base_consolidada = base_consolidada[columnas_finales]

    base_consolidada = convertir_tipos_base(base_consolidada)

    reporte_ventas = generar_reporte_ventas_por_proyecto(base_consolidada)

    ruta_excel = SALIDA_LISTA / "Base_Cobranzas_Consolidada.xlsx"

    ruta_reporte_ventas_excel = SALIDA_LISTA / "Reporte_Ventas_Por_Proyecto.xlsx"

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        base_consolidada.to_excel(
            writer,
            sheet_name="Base_Cobranzas",
            index=False
        )

    aplicar_formatos_excel(
        ruta_excel=ruta_excel,
        nombre_hoja="Base_Cobranzas"
    )


    with pd.ExcelWriter(ruta_reporte_ventas_excel, engine="openpyxl") as writer:
        reporte_ventas.to_excel(
            writer,
            sheet_name="Ventas_Proyecto",
            index=False
        )

    aplicar_formatos_reporte_ventas(
        ruta_excel=ruta_reporte_ventas_excel,
        nombre_hoja="Ventas_Proyecto"
    )

    print("====================================")
    print("CONSOLIDACIÓN TERMINADA")
    print("====================================")
    print("✅ Base de datos generado con exito:")
    print("✅ Reporte ventas generado con exito:")

    return base_consolidada, reporte_ventas


def main():
    if EJECUTAR_DESCARGA:
        ejecutar_descarga_cobranzas()

    consolidar_cobranzas()


if __name__ == "__main__":
    main()