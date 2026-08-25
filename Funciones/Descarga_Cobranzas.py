# Programa para descargar la base de datos de cobranzas en Evolta

# Creado por Eduardo Miguel Huamani Acosta                            15/07/26


# Este automatización trata de extraer datos desde Evolta, la limitación es que solo permite extraer información con un periodo 
# de un año, por lo que al momento de descargar de forma manual se podría cometer errores humanos

# Este código permite extraer información bajo el criterio de extraer la información por año y juntarlo en uno solo para 
# los analisis, de lo que se demoraba de 1 - 2 horas a 30 minutos aproximadamente

# =============================================================================================================================

# LIBRERIAS

from pathlib import Path
from datetime import datetime, date
import time
import shutil
import os
import sys
import re
import unicodedata

import pandas as pd
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================================================
# 1. CONFIGURACIÓN GENERAL
# =========================================================

BASE_DIR = Path(__file__).resolve().parent.parent    

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from Conexiones.connection import (FECHA_INICIO_REPORTE, FECHA_FIN_REPORTE, FECHA_CORTE_REPORTE,)

LOGIN_URL = "https://v4.evolta.pe/Login/Acceso/Index"
REPORTE_COBRANZA_URL = "https://v4.evolta.pe/Reportes/ReporteCobranza/Index"

DOWNLOADS_DIR = Path(r"C:\Users\ehuamani\Downloads")   #URL de la carpeta

load_dotenv(BASE_DIR / ".env")
load_dotenv(Path(__file__).resolve().parent / ".env")

LOGIN_AUTOMATICO = True

USUARIO_EVOLTA = os.getenv("USUARIO_EVOLTA")
CLAVE_EVOLTA = os.getenv("CLAVE_EVOLTA")

ENTRADA_COBRANZAS = BASE_DIR / "Flujo" / "Input" / "Cobranzas"
SALIDA_LISTA = BASE_DIR / "Flujo" / "Output"

ENTRADA_COBRANZAS.mkdir(parents=True, exist_ok=True)
SALIDA_LISTA.mkdir(parents=True, exist_ok=True)

FORMATO_EVOLTA = "Csv"
SEPARADOR_CSV = ","

CONVERTIR_CSV_A_EXCEL = True
OVERWRITE_EXISTING = True
BORRAR_ORIGINAL_DOWNLOADS = False

TIEMPO_ESPERA_DESCARGA = 180
INTENTOS_DESCARGA = 2

NOMBRE_HOJA_BASE = "Base_Cobranzas"
NOMBRE_BASE_CONSOLIDADA = "Base_Cobranzas_Consolidada.xlsx"
NOMBRE_BASE_AJUSTADA = "Base_Cobranzas_Ajustado.xlsx"

# Archivo maestro para obtener las áreas de cada inmueble
NOMBRE_ARCHIVO_UNIDADES = "Unidades-Inmobiliarias.xlsx"
NOMBRE_HOJA_UNIDADES = "Stock Comercial"

NOMBRE_HOJA_BASE_ORIGINAL_AJUSTADA = "Base_Original"
NOMBRE_HOJA_BASE_AJUSTADA = "Base_Ajustada"

COLUMNAS_FECHA = [
    "Fecha_Programada",
    "FechaPago",
    "FechaOperacionComercial",
]

COLUMNAS_NUMERO_FIJAS = [
    "Monto_Cuota",
    "Monto_Cuota_Pagado",
    "SaldoPorPagarCuota",
]

COLUMNAS_NUMERO_CON_INDICE = [
    "PrecioLista",
    "PrecioVenta",
    "AreaTechada",
    "AreaLibre",
]

EXTENSIONES_DESCARGA = [".csv", ".xlsx", ".xls"]


# =========================================================
# 2. UTILIDADES GENERALES
# =========================================================

# Limpia textos para compararlos mejor
def normalizar_texto(texto):
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )

# Normaliza el código del inmueble para poder cruzarlo entre bases
# Ejemplo: "E 14 - D 11" -> "E14D11"
def normalizar_codigo_inmueble(codigo):
    if pd.isna(codigo):
        return ""

    codigo = normalizar_texto(codigo)
    return re.sub(r"[^A-Z0-9]", "", codigo)

# Convierte la fecha de texto a numero
def convertir_fecha(fecha_texto: str) -> date:
    return datetime.strptime(fecha_texto, "%d/%m/%Y").date()

# Convierte la fecha de numero a texto 
def fecha_texto(fecha: date) -> str:
    return fecha.strftime("%d/%m/%Y")

# Lee la fecha de corte
def obtener_fecha_cierre_ajuste():
    if FECHA_CORTE_REPORTE is None or str(FECHA_CORTE_REPORTE).strip() == "":
        return None

    fecha_cierre = pd.to_datetime(
        FECHA_CORTE_REPORTE,
        errors="coerce",
        dayfirst=True,
    )

    if pd.isna(fecha_cierre):
        raise ValueError(
            "FECHA_CORTE_REPORTE no tiene un formato válido. "
            "Usa un formato como '30/06/2026'."
        )

    return fecha_cierre.normalize()

# Agarra la fecha inicial y final para que el reporte salga por años
def generar_rangos_por_anio(fecha_inicio_txt: str, fecha_fin_txt: str):
    inicio = convertir_fecha(fecha_inicio_txt)
    fin = convertir_fecha(fecha_fin_txt)

    if inicio > fin:
        raise ValueError("La fecha de inicio no puede ser mayor que la fecha de fin.")

    rangos = []

    for anio in range(inicio.year, fin.year + 1):
        desde = max(inicio, date(anio, 1, 1))
        hasta = min(fin, date(anio, 12, 31))

        rangos.append({
            "anio": anio,
            "fecha_inicio": fecha_texto(desde),
            "fecha_fin": fecha_texto(hasta),
        })

    return rangos

# Limpia una columna de montos y las convierte a numero
def convertir_columna_numero(serie):
    serie = serie.astype(str).str.strip()
    serie = serie.str.replace("S/", "", regex=False)
    serie = serie.str.replace("US$", "", regex=False)
    serie = serie.str.replace("$", "", regex=False)
    serie = serie.str.replace(" ", "", regex=False)
    serie = serie.str.replace(",", "", regex=False)
    serie = serie.replace("", pd.NA)

    return pd.to_numeric(serie, errors="coerce")

# Indica si una columna debe convertirse a numero
def es_columna_numero(columna):
    if columna in COLUMNAS_NUMERO_FIJAS:
        return True

    return any(
        re.match(rf"^{campo}_(\d+)$", str(columna))
        for campo in COLUMNAS_NUMERO_CON_INDICE
    )

# Aplica las conversiones a toda la base
def convertir_tipos_base(base):
    base = base.copy()

    for columna in COLUMNAS_FECHA:
        if columna in base.columns:
            base[columna] = pd.to_datetime(
                base[columna],
                errors="coerce",
                dayfirst=True,
            )

    for columna in base.columns:
        if es_columna_numero(columna):
            base[columna] = convertir_columna_numero(base[columna])

    return base


def validar_columnas(base, columnas, contexto):
    faltantes = [col for col in columnas if col not in base.columns]

    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas para {contexto}:\n"
            + "\n".join(f"- {col}" for col in faltantes)
        )

# Sirve el año desde el nombre del archivo 
def obtener_anio_archivo(nombre_archivo):
    coincidencia = re.search(r"(20\d{2})", nombre_archivo)
    return coincidencia.group(1) if coincidencia else ""


# =========================================================
# 3. SELENIUM / EVOLTA
# =========================================================

# Encuentra el boton Exportar de la página de Evolta
def obtener_xpath_exportar():
    return """
    //button[contains(normalize-space(.), 'Exportar')]
    | //a[contains(normalize-space(.), 'Exportar')]
    | //input[@value='Exportar']
    | //*[@title='Exportar']
    """

# Espera a que un elemento aparezca en la página
def esperar_elemento(driver, xpath, timeout=30):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )

# Espera que cargue la página de reporte de cobranzas
def esperar_pagina_reporte(driver, timeout=40):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, obtener_xpath_exportar()))
    )

# Lleva al navegador directamento al Reporte de Cobranzas de Evolta
def ir_a_reporte_cobranzas(driver):
    driver.get(REPORTE_COBRANZA_URL)
    esperar_pagina_reporte(driver)


def click_seguro(driver, elemento):
    try:
        elemento.click()
    except Exception:
        try:
            ActionChains(driver).move_to_element(elemento).pause(0.5).click().perform()
        except Exception:
            driver.execute_script("arguments[0].click();", elemento)

# Inicia sesión en Evolta
def iniciar_sesion_evolta(driver):
    print("\n💻 Ingresando a Evolta automáticamente...")

    if not USUARIO_EVOLTA or not CLAVE_EVOLTA:
        raise ValueError(
            "No se encontraron USUARIO_EVOLTA o CLAVE_EVOLTA en el archivo .env."
        )

    driver.get(LOGIN_URL)

    xpath_usuario = """
    //input[@type='email']
    | //input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'usuario')]
    | //input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'usuario')]
    | //input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'usuario')]
    | //input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]
    | //input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]
    | //input[not(@type='password') and not(@type='hidden')][1]
    """

    xpath_clave = """
    //input[@type='password']
    | //input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'clave')]
    | //input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'clave')]
    | //input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'clave')]
    | //input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password')]
    | //input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password')]
    """

    xpath_boton_ingresar = """
    //button[contains(normalize-space(.), 'Ingresar')]
    | //button[contains(normalize-space(.), 'Entrar')]
    | //button[contains(normalize-space(.), 'Acceder')]
    | //input[@type='submit']
    | //button[@type='submit']
    """

    campo_usuario = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, xpath_usuario))
    )
    campo_clave = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, xpath_clave))
    )

    for campo, valor in [
        (campo_usuario, USUARIO_EVOLTA),
        (campo_clave, CLAVE_EVOLTA),
    ]:
        campo.click()
        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(Keys.BACKSPACE)
        campo.send_keys(valor)

    boton_ingresar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, xpath_boton_ingresar))
    )

    click_seguro(driver, boton_ingresar)
    time.sleep(3)

    ir_a_reporte_cobranzas(driver)

    print("\n✅ Sesión iniciada correctamente.")

# Selecciona el formato de descarga de Evolta
def seleccionar_formato(driver, formato: str):
    if formato not in ["Csv", "Excel"]:
        raise ValueError("El formato debe ser 'Csv' o 'Excel'.")

    xpath_radio = f"//*[normalize-space()='{formato}']/preceding::input[@type='radio'][1]"

    radio = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, xpath_radio))
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", radio)
    time.sleep(0.5)

# Escribe la fecha en Evolta
def escribir_fecha(driver, nombre_campo: str, fecha: str):
    print(f"Escribiendo {nombre_campo}: {fecha}")

    xpath = f"//*[normalize-space()='{nombre_campo}']/following::input[1]"
    campo = esperar_elemento(driver, xpath, timeout=30)

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
    time.sleep(0.5)

    campo.click()
    time.sleep(0.2)
    campo.send_keys(Keys.CONTROL, "a")
    campo.send_keys(Keys.BACKSPACE)
    campo.send_keys(fecha)
    campo.send_keys(Keys.TAB)

    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        arguments[0].blur();
    """, campo)

    time.sleep(1)

# Hace clic Exportar desde Evolta
def click_exportar(driver):
    print("Buscando botón Exportar...")

    boton = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, obtener_xpath_exportar()))
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
    time.sleep(1)

    print("Haciendo clic en Exportar...")
    click_seguro(driver, boton)
    time.sleep(3)

# Configura Chrome para que descargue los archivos en la carpeta indicada
def configurar_descargas_chrome(driver):
    ruta_descarga = str(DOWNLOADS_DIR.resolve())
    parametros = {
        "behavior": "allow",
        "downloadPath": ruta_descarga,
    }

    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", parametros)
    except Exception:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", parametros)


def crear_driver_chrome():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": str(DOWNLOADS_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }

    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    configurar_descargas_chrome(driver)

    return driver


# =========================================================
# 4. DESCARGAS
# =========================================================

# Lee que archivos existen en la carpeta Descargas antes de descargar.
def obtener_archivos_downloads():
    return {
        archivo.name
        for archivo in DOWNLOADS_DIR.iterdir()
        if archivo.is_file()
        and archivo.suffix.lower() in EXTENSIONES_DESCARGA
    }

# Espera que la descarga termine
def esperar_nueva_descarga_downloads(archivos_antes, tiempo_inicio, timeout=180):
    print("⌛ Esperando archivo nuevo en Descargas...")

    tiempo_limite = time.time() + timeout
    ultimo_aviso = time.time()

    while time.time() < tiempo_limite:
        temporales_recientes = [
            temporal
            for temporal in DOWNLOADS_DIR.glob("*.crdownload")
            if temporal.exists()
            and temporal.stat().st_mtime >= tiempo_inicio - 2
        ]

        if temporales_recientes:
            time.sleep(1)
            continue

        candidatos = []

        for archivo in DOWNLOADS_DIR.iterdir():
            if not archivo.is_file() or archivo.suffix.lower() not in EXTENSIONES_DESCARGA:
                continue

            try:
                nombre_nuevo = archivo.name not in archivos_antes
                modificado_reciente = archivo.stat().st_mtime >= tiempo_inicio - 2
            except Exception:
                continue

            if nombre_nuevo or modificado_reciente:
                candidatos.append(archivo)

        if candidatos:
            archivo_nuevo = max(candidatos, key=lambda f: f.stat().st_mtime)

            try:
                tamaño_1 = archivo_nuevo.stat().st_size
                time.sleep(1)
                tamaño_2 = archivo_nuevo.stat().st_size

                if tamaño_1 == tamaño_2 and tamaño_2 > 0:
                    return archivo_nuevo
            except Exception:
                pass

        if time.time() - ultimo_aviso >= 15:
            print("Todavía esperando descarga...")
            ultimo_aviso = time.time()

        time.sleep(1)

    raise TimeoutError("No se detectó ningún archivo nuevo en Descargas.")

# Borra archivos anteriores del mismo año antes de descargar
def limpiar_archivos_previos_anio(anio: int):
    if not OVERWRITE_EXISTING:
        return

    for carpeta in [DOWNLOADS_DIR, ENTRADA_COBRANZAS]:
        for extension in EXTENSIONES_DESCARGA:
            archivo = carpeta / f"Cobranzas_{anio}{extension}"

            if archivo.exists():
                try:
                    archivo.unlink()
                except Exception:
                    pass

# Renombra a los archivos descargados
def renombrar_en_downloads(archivo_descargado: Path, anio: int):
    extension = archivo_descargado.suffix.lower()

    if extension not in EXTENSIONES_DESCARGA:
        raise ValueError(f"Extensión no esperada: {extension}")

    destino = DOWNLOADS_DIR / f"Cobranzas_{anio}{extension}"

    if destino.exists() and OVERWRITE_EXISTING:
        destino.unlink()

    if destino.exists() and not OVERWRITE_EXISTING:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = DOWNLOADS_DIR / f"Cobranzas_{anio}_{timestamp}{extension}"

    archivo_descargado.rename(destino)

    if not destino.exists():
        raise FileNotFoundError(f"No se renombró correctamente en Descargas: {destino}")

    return destino

# Copia el archivo de Descargas hacia la carpeta del proyecto
def copiar_renombrado_a_cobranzas(archivo_renombrado: Path):
    destino = ENTRADA_COBRANZAS / archivo_renombrado.name

    if destino.exists() and OVERWRITE_EXISTING:
        destino.unlink()

    if destino.exists() and not OVERWRITE_EXISTING:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = ENTRADA_COBRANZAS / f"{archivo_renombrado.stem}_{timestamp}{archivo_renombrado.suffix}"

    shutil.copy2(archivo_renombrado, destino)

    if not destino.exists():
        raise FileNotFoundError(f"No se copió correctamente el archivo a: {destino}")

    if BORRAR_ORIGINAL_DOWNLOADS:
        archivo_renombrado.unlink()
        print("Archivo original eliminado de Descargas.")

    return destino

# Descarga un año en especifico desde Evolta
def descargar_rango(driver, anio, fecha_inicio, fecha_fin):
    ultimo_error = None

    for intento in range(1, INTENTOS_DESCARGA + 1):
        print(f"Intento de descarga {intento} de {INTENTOS_DESCARGA}")

        try:
            archivos_antes = obtener_archivos_downloads()

            seleccionar_formato(driver, FORMATO_EVOLTA)
            escribir_fecha(driver, "Fecha de inicio", fecha_inicio)
            escribir_fecha(driver, "Fecha de fin", fecha_fin)

            tiempo_inicio_descarga = time.time()
            click_exportar(driver)

            return esperar_nueva_descarga_downloads(
                archivos_antes=archivos_antes,
                tiempo_inicio=tiempo_inicio_descarga,
                timeout=TIEMPO_ESPERA_DESCARGA,
            )

        except Exception as e:
            ultimo_error = e
            print(f"No se detectó descarga en el intento {intento}.")
            print(e)

            if intento < INTENTOS_DESCARGA:
                print("Recargando página de reporte para intentar nuevamente...")
                ir_a_reporte_cobranzas(driver)
                time.sleep(2)

    raise ultimo_error


def procesar_descarga_anual(driver, rango):
    anio = rango["anio"]
    fecha_inicio = rango["fecha_inicio"]
    fecha_fin = rango["fecha_fin"]

    print("\n====================================")
    print(f"PROCESANDO AÑO {anio}")
    print(f"Fecha inicio: {fecha_inicio}")
    print(f"Fecha fin: {fecha_fin}")
    print("====================================")

    limpiar_archivos_previos_anio(anio)

    archivo_descargado = descargar_rango(
        driver=driver,
        anio=anio,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    archivo_renombrado = renombrar_en_downloads(
        archivo_descargado=archivo_descargado,
        anio=anio,
    )

    archivo_final = copiar_renombrado_a_cobranzas(archivo_renombrado)

    if CONVERTIR_CSV_A_EXCEL and archivo_final.suffix.lower() == ".csv":
        convertir_csv_a_excel_seguro(archivo_final)

    print(f"Año {anio} descargado correctamente.")

    return archivo_final


# =========================================================
# 5. LECTURA, CONVERSIÓN Y CONSOLIDACIÓN
# =========================================================

# Lee el CSV de forma segura
def leer_csv_seguro(ruta_csv: Path) -> pd.DataFrame:
    for encoding in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
        try:
            return pd.read_csv(
                ruta_csv,
                sep=SEPARADOR_CSV,
                dtype=str,
                encoding=encoding,
                keep_default_na=False,
            )
        except Exception:
            continue

    raise ValueError(f"No se pudo leer el CSV: {ruta_csv}")

# Convierte CSV a Excel
def convertir_csv_a_excel_seguro(ruta_csv: Path):
    df = leer_csv_seguro(ruta_csv)
    ruta_excel = ruta_csv.with_suffix(".xlsx")

    if ruta_excel.exists() and OVERWRITE_EXISTING:
        ruta_excel.unlink()

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_CSV", index=False)
        ws = writer.book["Base_CSV"]

        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = "@"

    print(f"✅ Excel generado: {ruta_csv.name}")

    return ruta_excel

# Lee un archivo de cobranzas
def leer_archivo_cobranzas(ruta_archivo: Path) -> pd.DataFrame:
    if ruta_archivo.suffix.lower() == ".csv":
        return leer_csv_seguro(ruta_archivo)

    if ruta_archivo.suffix.lower() in [".xlsx", ".xls"]:
        return pd.read_excel(
            ruta_archivo,
            dtype=str,
            keep_default_na=False,
        )

    raise ValueError(f"Formato no permitido para consolidar: {ruta_archivo}")

# Busca los archivos que se van a consolidar
def obtener_archivos_para_consolidar():
    archivos_csv = sorted(ENTRADA_COBRANZAS.glob("Cobranzas_*.csv"))

    if archivos_csv:
        return archivos_csv

    archivos_excel = sorted(ENTRADA_COBRANZAS.glob("Cobranzas_*.xlsx"))

    if archivos_excel:
        return archivos_excel

    raise FileNotFoundError(
        f"❌ No se encontraron archivos Cobranzas_*.csv ni Cobranzas_*.xlsx en: {ENTRADA_COBRANZAS}"
    )

# Revisa todos los archivs y arma una lista con todas las columnas existentes
def obtener_columnas_union(archivos_cobranzas):
    columnas_union = []
    columnas_vistas = set()

    for archivo in archivos_cobranzas:
        df = leer_archivo_cobranzas(archivo)

        for columna in df.columns:
            if columna not in columnas_vistas:
                columnas_union.append(columna)
                columnas_vistas.add(columna)

    return columnas_union

# Ordena las columnas de la base final
def ordenar_columnas_cobranzas(columnas_union):
    campos_inmueble = [
        "Tipo_Construccion",
        "Nombre_Tipo_Construccion",
        "TipoInmueble",
        "NroInmueble",
        "AreaTechada",
        "AreaLibre",
        "PrecioLista",
        "PrecioVenta",
    ]

    columnas_auxiliares = [
        "Archivo_Origen",
        "Anio_Descarga",
        "Fecha_Inicio_Reporte",
        "Fecha_Fin_Reporte",
    ]

    indices = set()
    columnas_a_reordenar = set()

    for columna in columnas_union:
        for campo in campos_inmueble:
            coincidencia = re.match(rf"^{campo}_(\d+)$", str(columna))

            if coincidencia:
                indices.add(int(coincidencia.group(1)))
                columnas_a_reordenar.add(columna)

    columnas_inmueble = []

    for indice in sorted(indices):
        for campo in campos_inmueble:
            columna = f"{campo}_{indice}"

            if columna in columnas_union:
                columnas_inmueble.append(columna)

    if "Nombres_Titular" in columnas_union:
        columnas_a_reordenar.add("Nombres_Titular")
        columnas_inmueble.append("Nombres_Titular")

    posiciones = [
        i
        for i, columna in enumerate(columnas_union)
        if columna in columnas_a_reordenar
    ]

    posicion_insertar = min(posiciones) if posiciones else len(columnas_union)

    columnas_sin_bloque = [
        columna
        for columna in columnas_union
        if columna not in columnas_a_reordenar
    ]

    columnas_reordenadas = (
        columnas_sin_bloque[:posicion_insertar]
        + columnas_inmueble
        + columnas_sin_bloque[posicion_insertar:]
    )

    columnas_finales = [
        columna
        for columna in columnas_reordenadas
        if columna not in columnas_auxiliares
    ]

    return columnas_finales + columnas_auxiliares


def agregar_columnas_auxiliares(df, archivo, columnas_finales):
    df = df.copy()
    anio_archivo = obtener_anio_archivo(archivo.name)

    df["Archivo_Origen"] = archivo.name
    df["Anio_Descarga"] = anio_archivo
    df["Fecha_Inicio_Reporte"] = FECHA_INICIO_REPORTE
    df["Fecha_Fin_Reporte"] = FECHA_FIN_REPORTE

    for columna in columnas_finales:
        if columna not in df.columns:
            df[columna] = ""

    return df[columnas_finales]



# Agrega el área techada y el área libre de cada inmueble desde Stock Comercial
def agregar_areas_inmuebles(base):
    ruta_unidades = BASE_DIR / "Flujo" / "Input" / NOMBRE_ARCHIVO_UNIDADES

    if not ruta_unidades.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de unidades inmobiliarias:\n{ruta_unidades}"
        )

    print("\n====================================")
    print("AGREGANDO ÁREAS DE INMUEBLES")
    print("====================================")
    print(f"Archivo fuente: {ruta_unidades.name}")
    print(f"Hoja fuente: {NOMBRE_HOJA_UNIDADES}")

    stock = pd.read_excel(
        ruta_unidades,
        sheet_name=NOMBRE_HOJA_UNIDADES,
        engine="openpyxl",
        dtype=str,
        keep_default_na=False,
    )

    validar_columnas(
        stock,
        [
            "Proyecto",
            "TipoInmueble",
            "NroInmuebleActual",
            "Areatechada",
            "AreaLibre",
        ],
        "cruce de áreas con Stock Comercial",
    )

    # Normaliza las llaves del archivo maestro.
    stock["_Proyecto_Cruce"] = stock["Proyecto"].apply(normalizar_texto)
    stock["_TipoInmueble_Cruce"] = stock["TipoInmueble"].apply(normalizar_texto)
    stock["_NroInmueble_Cruce"] = stock["NroInmuebleActual"].apply(
        normalizar_codigo_inmueble
    )

    # Las áreas deben quedar como valores numéricos.
    stock["Areatechada"] = convertir_columna_numero(stock["Areatechada"])
    stock["AreaLibre"] = convertir_columna_numero(stock["AreaLibre"])

    columnas_llave = [
        "_Proyecto_Cruce",
        "_TipoInmueble_Cruce",
        "_NroInmueble_Cruce",
    ]

    # Si Stock Comercial tiene la misma unidad repetida, conserva una sola fila.
    # Antes se valida que esas repeticiones no tengan áreas diferentes.
    duplicados = stock.loc[
        stock.duplicated(subset=columnas_llave, keep=False)
        & (stock["_NroInmueble_Cruce"] != "")
    ].copy()

    if not duplicados.empty:
        conflictos = (
            duplicados.groupby(columnas_llave, dropna=False)[
                ["Areatechada", "AreaLibre"]
            ]
            .nunique(dropna=False)
        )

        conflictos = conflictos[
            (conflictos["Areatechada"] > 1)
            | (conflictos["AreaLibre"] > 1)
        ]

        if not conflictos.empty:
            raise ValueError(
                "Se encontraron unidades repetidas en Stock Comercial con "
                "áreas diferentes. Revisa Proyecto + TipoInmueble + "
                "NroInmuebleActual antes de continuar."
            )

    stock_lookup = stock.drop_duplicates(
        subset=columnas_llave,
        keep="first",
    ).copy()

    mapa_area_techada = stock_lookup.set_index(columnas_llave)[
        "Areatechada"
    ].to_dict()

    mapa_area_libre = stock_lookup.set_index(columnas_llave)[
        "AreaLibre"
    ].to_dict()

    base_area = base.copy()

    validar_columnas(
        base_area,
        ["Proyecto"],
        "cruce de áreas en la base de cobranzas",
    )

    proyecto_normalizado = base_area["Proyecto"].apply(normalizar_texto)

    # Detecta automáticamente NroInmueble_1, NroInmueble_2, etc.
    indices = []

    for columna in base_area.columns:
        coincidencia = re.match(r"^NroInmueble_(\d+)$", str(columna))

        if coincidencia:
            indices.append(int(coincidencia.group(1)))

    indices = sorted(set(indices))

    if not indices:
        raise ValueError(
            "No se encontraron columnas NroInmueble_1, NroInmueble_2, etc. "
            "para realizar el cruce de áreas."
        )

    total_con_inmueble = 0
    total_encontrados = 0

    for indice in indices:
        columna_tipo = f"TipoInmueble_{indice}"
        columna_inmueble = f"NroInmueble_{indice}"
        columna_area_techada = f"AreaTechada_{indice}"
        columna_area_libre = f"AreaLibre_{indice}"

        if columna_tipo not in base_area.columns:
            print(
                f"⚠️ No existe {columna_tipo}; se omite el cruce del inmueble {indice}."
            )
            continue

        tipo_normalizado = base_area[columna_tipo].apply(normalizar_texto)
        inmueble_normalizado = base_area[columna_inmueble].apply(
            normalizar_codigo_inmueble
        )

        llaves = list(
            zip(
                proyecto_normalizado,
                tipo_normalizado,
                inmueble_normalizado,
            )
        )

        base_area[columna_area_techada] = [
            mapa_area_techada.get(llave, pd.NA)
            if llave[2] != ""
            else pd.NA
            for llave in llaves
        ]

        base_area[columna_area_libre] = [
            mapa_area_libre.get(llave, pd.NA)
            if llave[2] != ""
            else pd.NA
            for llave in llaves
        ]

        con_inmueble = inmueble_normalizado.ne("")
        encontrados = (
            con_inmueble
            & (
                base_area[columna_area_techada].notna()
                | base_area[columna_area_libre].notna()
            )
        )

        cantidad_con_inmueble = int(con_inmueble.sum())
        cantidad_encontrados = int(encontrados.sum())
        cantidad_no_encontrados = cantidad_con_inmueble - cantidad_encontrados

        total_con_inmueble += cantidad_con_inmueble
        total_encontrados += cantidad_encontrados

        print(
            f"Inmueble {indice}: "
            f"{cantidad_encontrados} de {cantidad_con_inmueble} encontrados"
        )

        if cantidad_no_encontrados > 0:
            print(
                f"⚠️ Inmueble {indice}: "
                f"{cantidad_no_encontrados} registros sin coincidencia de área."
            )

    print("------------------------------------")
    print(
        f"Total de inmuebles con área encontrada: "
        f"{total_encontrados} de {total_con_inmueble}"
    )

    return base_area


def formatear_hoja_base(ws):
    encabezados = {
        celda.value: celda.column
        for celda in ws[1]
    }

    for columna in COLUMNAS_FECHA:
        if columna in encabezados:
            col_idx = encabezados[columna]

            for fila in range(2, ws.max_row + 1):
                ws.cell(row=fila, column=col_idx).number_format = "dd/mm/yyyy"

    for columna, col_idx in encabezados.items():
        if es_columna_numero(columna):
            for fila in range(2, ws.max_row + 1):
                ws.cell(row=fila, column=col_idx).number_format = "#,##0.00"

    ws.freeze_panes = "A2"

# Funcion principal del proceso, arma la base de datos
def generar_base_cobranzas_consolidada():
    archivos_cobranzas = obtener_archivos_para_consolidar()

    print("\n✅ Archivos que se van a consolidar:")
    for archivo in archivos_cobranzas:
        print("-", archivo.name)

    columnas_union = obtener_columnas_union(archivos_cobranzas)
    columnas_finales = ordenar_columnas_cobranzas(columnas_union)

    bases = []

    for archivo in archivos_cobranzas:
        df = leer_archivo_cobranzas(archivo)
        df = agregar_columnas_auxiliares(df, archivo, columnas_finales)
        bases.append(df)

        print(f"{archivo.name}: {len(df)} filas, {len(df.columns)} columnas")

    base_consolidada = pd.concat(bases, ignore_index=True)
    base_consolidada = base_consolidada[columnas_finales]
    base_consolidada = convertir_tipos_base(base_consolidada)

    # Agrega AreaTechada_n y AreaLibre_n desde Unidades-Inmobiliarias.xlsx
    base_consolidada = agregar_areas_inmuebles(base_consolidada)

    # Reordena las columnas para que las áreas queden junto al inmueble correspondiente
    columnas_finales = ordenar_columnas_cobranzas(
        list(base_consolidada.columns)
    )
    base_consolidada = base_consolidada[columnas_finales]
    base_consolidada = convertir_tipos_base(base_consolidada)

    ruta_base = SALIDA_LISTA / NOMBRE_BASE_CONSOLIDADA

    with pd.ExcelWriter(ruta_base, engine="openpyxl") as writer:
        base_consolidada.to_excel(
            writer,
            sheet_name=NOMBRE_HOJA_BASE,
            index=False,
        )

        formatear_hoja_base(writer.book[NOMBRE_HOJA_BASE])

    print("\n✅ Base consolidada original generada correctamente:")
    print(ruta_base.resolve())

    return ruta_base


# =========================================================
# 6. AJUSTES DE BASE
# =========================================================

# Busca que columna usar para identificar al cliente cuando se ajustan saldos negativos
def obtener_columna_cliente_ajuste(base):
    if "NroDocumento" in base.columns:
        return "NroDocumento"

    if "Nombres_Titular" in base.columns:
        return "Nombres_Titular"

    raise ValueError(
        "No se encontró una columna válida para identificar al cliente "
        "(NroDocumento o Nombres_Titular)."
    )

# Busca q columna usar para identificar la unidad inmobiliaria
def obtener_columna_departamento_ajuste(base):
    if "NroInmueble_1" in base.columns:
        return "NroInmueble_1"

    if "NroInmueble" in base.columns:
        return "NroInmueble"

    raise ValueError(
        "No se encontró una columna válida para identificar el departamento "
        "(NroInmueble_1 o NroInmueble)."
    )

# Cambia el estado de las cuotas según la fecha de corte
def reclasificar_estado_por_fecha_cierre(base, fecha_cierre):
    base_estado = base.copy()

    if fecha_cierre is None:
        return base_estado, 0, 0

    validar_columnas(
        base_estado,
        ["Estado", "Fecha_Programada", "SaldoPorPagarCuota"],
        "reclasificar estados",
    )

    base_estado["Fecha_Programada"] = pd.to_datetime(
        base_estado["Fecha_Programada"],
        errors="coerce",
        dayfirst=True,
    )

    base_estado["SaldoPorPagarCuota"] = pd.to_numeric(
        base_estado["SaldoPorPagarCuota"],
        errors="coerce",
    ).fillna(0)

    condicion_pago_posterior_lejano = pd.Series(False, index=base_estado.index)

    if "FechaPago" in base_estado.columns and "Monto_Cuota_Pagado" in base_estado.columns:
        fecha_pago = pd.to_datetime(
            base_estado["FechaPago"],
            errors="coerce",
            dayfirst=True,
        )

        monto_pagado = pd.to_numeric(
            base_estado["Monto_Cuota_Pagado"],
            errors="coerce",
        ).fillna(0)

        inicio_mes_siguiente = pd.Timestamp(
            year=fecha_cierre.year,
            month=fecha_cierre.month,
            day=1,
        ) + pd.DateOffset(months=1)

        inicio_mes_posterior = inicio_mes_siguiente + pd.DateOffset(months=1)

        condicion_pago_posterior_lejano = (
            fecha_pago.notna()
            & (fecha_pago >= inicio_mes_posterior)
            & (monto_pagado != 0)
        )

    condicion_con_saldo = (
        (base_estado["SaldoPorPagarCuota"] > 0)
        & base_estado["Fecha_Programada"].notna()
        & (~condicion_pago_posterior_lejano)
    )

    condicion_vencido = condicion_con_saldo & (base_estado["Fecha_Programada"] <= fecha_cierre)
    condicion_pendiente = condicion_con_saldo & (base_estado["Fecha_Programada"] > fecha_cierre)

    base_estado.loc[condicion_vencido, "Estado"] = "VENCIDO"
    base_estado.loc[condicion_pendiente, "Estado"] = "PENDIENTE"

    return base_estado, int(condicion_vencido.sum()), int(condicion_pendiente.sum())
# Aplica la logica del cierre mensual 
def aplicar_fecha_cierre_base_ajustada(base):
    fecha_cierre = obtener_fecha_cierre_ajuste()

    if fecha_cierre is None:
        print("ℹ️ No se aplicó fecha de cierre porque FECHA_CORTE_REPORTE está vacío.")
        return base.copy()

    validar_columnas(
        base,
        [
            "Estado",
            "Fecha_Programada",
            "FechaPago",
            "Monto_Cuota_Pagado",
            "SaldoPorPagarCuota",
        ],
        "aplicar fecha de cierre",
    )

    base_cierre = base.copy()

    for columna in ["Fecha_Programada", "FechaPago"]:
        base_cierre[columna] = pd.to_datetime(
            base_cierre[columna],
            errors="coerce",
            dayfirst=True,
        )

    for columna in ["Monto_Cuota_Pagado", "SaldoPorPagarCuota"]:
        base_cierre[columna] = pd.to_numeric(
            base_cierre[columna],
            errors="coerce",
        ).fillna(0)

    if "Monto_Cuota" in base_cierre.columns:
        base_cierre["Monto_Cuota"] = pd.to_numeric(
            base_cierre["Monto_Cuota"],
            errors="coerce",
        ).fillna(0)

    inicio_mes_siguiente = pd.Timestamp(
        year=fecha_cierre.year,
        month=fecha_cierre.month,
        day=1,
    ) + pd.DateOffset(months=1)

    inicio_mes_posterior = inicio_mes_siguiente + pd.DateOffset(months=1)

    condicion_pago_mes_siguiente = (
        (base_cierre["FechaPago"].notna())
        & (base_cierre["FechaPago"] >= inicio_mes_siguiente)
        & (base_cierre["FechaPago"] < inicio_mes_posterior)
        & (base_cierre["Monto_Cuota_Pagado"] != 0)
    )

    condicion_pago_posterior_lejano = (
        (base_cierre["FechaPago"].notna())
        & (base_cierre["FechaPago"] >= inicio_mes_posterior)
        & (base_cierre["Monto_Cuota_Pagado"] != 0)
    )

    cantidad_pagos_posteriores = int(condicion_pago_mes_siguiente.sum())
    cantidad_pagos_posteriores_lejanos = int(condicion_pago_posterior_lejano.sum())

    if cantidad_pagos_posteriores > 0:
        if "Monto_Cuota" in base_cierre.columns:
            saldo_restaurado = base_cierre["Monto_Cuota"].copy().fillna(0)
            saldo_restaurado = saldo_restaurado.mask(saldo_restaurado < 0, 0)
        else:
            saldo_restaurado = base_cierre["SaldoPorPagarCuota"].copy().fillna(0)

        base_cierre.loc[condicion_pago_mes_siguiente, "Monto_Cuota_Pagado"] = 0
        base_cierre.loc[condicion_pago_mes_siguiente, "SaldoPorPagarCuota"] = saldo_restaurado.loc[condicion_pago_mes_siguiente]

    base_cierre, cantidad_vencidos, cantidad_pendientes = reclasificar_estado_por_fecha_cierre(
        base_cierre,
        fecha_cierre,
    )

    print("\n====================================")
    print("AJUSTE POR FECHA DE CIERRE")
    print("====================================")
    print(f"Fecha de cierre: {fecha_cierre.strftime('%d/%m/%Y')}")
    print(f"Pagos del mes siguiente al cierre revertidos en Base_Ajustada: {cantidad_pagos_posteriores}")
    print(f"Pagos posteriores lejanos conservados para revisión: {cantidad_pagos_posteriores_lejanos}")
    print(f"Filas con saldo reclasificadas como VENCIDO: {cantidad_vencidos}")
    print(f"Filas con saldo reclasificadas como PENDIENTE: {cantidad_pendientes}")

    return base_cierre
# Reestructura los saldos negativos
def aplicar_ajuste_saldos_negativos(base):
    validar_columnas(
        base,
        ["Proyecto", "Fecha_Programada", "SaldoPorPagarCuota", "Monto_Cuota_Pagado"],
        "ajuste de saldos negativos",
    )

    columna_cliente = obtener_columna_cliente_ajuste(base)
    columna_departamento = obtener_columna_departamento_ajuste(base)

    base_ajustada = base.copy()

    base_ajustada["Fecha_Programada"] = pd.to_datetime(
        base_ajustada["Fecha_Programada"],
        errors="coerce",
        dayfirst=True,
    )

    base_ajustada["SaldoPorPagarCuota"] = pd.to_numeric(
        base_ajustada["SaldoPorPagarCuota"],
        errors="coerce",
    ).fillna(0)

    base_ajustada["_Orden_Original_Ajuste"] = range(len(base_ajustada))

    columnas_grupo = ["Proyecto", columna_cliente, columna_departamento]

    total_negativos = int((base_ajustada["SaldoPorPagarCuota"] < 0).sum())
    negativos_aplicados = 0
    monto_reestructurado = 0.0

    llave_completa = base_ajustada[columnas_grupo].notna().all(axis=1)

    for columna in columnas_grupo:
        llave_completa &= base_ajustada[columna].astype(str).str.strip() != ""

    grupos = base_ajustada.loc[llave_completa].groupby(
        columnas_grupo,
        dropna=False,
        sort=False,
    )

    for _, grupo in grupos:
        grupo_ordenado = grupo.sort_values(
            by=["Fecha_Programada", "_Orden_Original_Ajuste"],
            ascending=[True, True],
        )

        indices_grupo = list(grupo_ordenado.index)

        for posicion, idx_negativo in enumerate(indices_grupo):
            saldo_negativo = float(base_ajustada.at[idx_negativo, "SaldoPorPagarCuota"])

            if saldo_negativo >= 0:
                continue

            credito_pendiente = abs(saldo_negativo)
            credito_inicial = credito_pendiente

            for idx_destino in indices_grupo[posicion + 1:]:
                if credito_pendiente <= 0:
                    break

                saldo_destino = float(base_ajustada.at[idx_destino, "SaldoPorPagarCuota"])

                if saldo_destino <= 0:
                    continue

                monto_a_aplicar = min(credito_pendiente, saldo_destino)

                base_ajustada.at[idx_destino, "SaldoPorPagarCuota"] = saldo_destino - monto_a_aplicar
                credito_pendiente -= monto_a_aplicar
                monto_reestructurado += monto_a_aplicar

            if credito_pendiente <= 0:
                base_ajustada.at[idx_negativo, "SaldoPorPagarCuota"] = 0
                negativos_aplicados += 1
            else:
                base_ajustada.at[idx_negativo, "SaldoPorPagarCuota"] = -credito_pendiente

                if credito_pendiente < credito_inicial:
                    negativos_aplicados += 1

    base_ajustada = base_ajustada.drop(columns=["_Orden_Original_Ajuste"])

    print("\n====================================")
    print("AJUSTE DE SALDOS NEGATIVOS")
    print("====================================")
    print(f"Saldos negativos encontrados: {total_negativos}")
    print(f"Filas negativas reestructuradas total/parcial: {negativos_aplicados}")
    print(f"Monto aplicado a cuotas futuras: {monto_reestructurado:,.2f}")

    return base_ajustada

# Marca cuotas parciales pagadas
def aplicar_estado_parcial_cancelado(base):
    fecha_cierre = obtener_fecha_cierre_ajuste()
    base_parcial = base.copy()

    validar_columnas(
        base_parcial,
        ["Estado", "Fecha_Programada", "Monto_Cuota_Pagado", "SaldoPorPagarCuota"],
        "marcar parciales",
    )

    base_parcial["Fecha_Programada"] = pd.to_datetime(
        base_parcial["Fecha_Programada"],
        errors="coerce",
        dayfirst=True,
    )

    for columna in ["Monto_Cuota_Pagado", "SaldoPorPagarCuota"]:
        base_parcial[columna] = pd.to_numeric(
            base_parcial[columna],
            errors="coerce",
        ).fillna(0)

    condicion_pago_posterior_lejano = pd.Series(False, index=base_parcial.index)

    if fecha_cierre is not None and "FechaPago" in base_parcial.columns:
        fecha_pago = pd.to_datetime(
            base_parcial["FechaPago"],
            errors="coerce",
            dayfirst=True,
        )

        inicio_mes_siguiente = pd.Timestamp(
            year=fecha_cierre.year,
            month=fecha_cierre.month,
            day=1,
        ) + pd.DateOffset(months=1)

        inicio_mes_posterior = inicio_mes_siguiente + pd.DateOffset(months=1)

        condicion_pago_posterior_lejano = (
            fecha_pago.notna()
            & (fecha_pago >= inicio_mes_posterior)
            & (base_parcial["Monto_Cuota_Pagado"] != 0)
        )

    condicion_parcial = (
        (base_parcial["Monto_Cuota_Pagado"] > 0)
        & (base_parcial["SaldoPorPagarCuota"] > 0)
        & (~condicion_pago_posterior_lejano)
    )

    if fecha_cierre is not None:
        condicion_vencido = (
            condicion_parcial
            & base_parcial["Fecha_Programada"].notna()
            & (base_parcial["Fecha_Programada"] <= fecha_cierre)
        )

        condicion_pendiente = (
            condicion_parcial
            & base_parcial["Fecha_Programada"].notna()
            & (base_parcial["Fecha_Programada"] > fecha_cierre)
        )
    else:
        estado_actual = base_parcial["Estado"].astype(str).apply(normalizar_texto)
        condicion_vencido = condicion_parcial & estado_actual.str.contains("VENCIDO", na=False)
        condicion_pendiente = condicion_parcial & estado_actual.str.contains("PENDIENTE", na=False)

    base_parcial.loc[condicion_vencido, "Estado"] = "VENCIDO - PARC. CANCELADO"
    base_parcial.loc[condicion_pendiente, "Estado"] = "PENDIENTE - PARC. CANCELADO"

    cantidad_parcial = int((condicion_vencido | condicion_pendiente).sum())

    print("\n====================================")
    print("ESTADO PARCIAL CANCELADO")
    print("====================================")
    print(f"Cuotas parcialmente canceladas marcadas: {cantidad_parcial}")

    return base_parcial
# Elimina venas posteriores al cierre
def filtrar_ventas_hasta_fecha_cierre(base):
    fecha_cierre = obtener_fecha_cierre_ajuste()

    if fecha_cierre is None:
        print("ℹ️ No se filtró FechaOperacionComercial porque FECHA_CORTE_REPORTE está vacío.")
        return base.copy()

    if "FechaOperacionComercial" not in base.columns:
        raise ValueError(
            "❌ No se puede filtrar la base ajustada porque no existe la columna "
            "FechaOperacionComercial."
        )

    base_filtrada = base.copy()

    fecha_operacion = pd.to_datetime(
        base_filtrada["FechaOperacionComercial"],
        errors="coerce",
        dayfirst=True,
    )

    condicion_mantener = fecha_operacion.notna() & (fecha_operacion <= fecha_cierre)

    filas_antes = len(base_filtrada)
    filas_eliminadas = int((~condicion_mantener).sum())
    filas_posteriores = int((fecha_operacion.notna() & (fecha_operacion > fecha_cierre)).sum())
    filas_sin_fecha = int(fecha_operacion.isna().sum())

    base_filtrada = base_filtrada.loc[condicion_mantener].copy()

    print("\n====================================")
    print("FILTRO DE VENTAS POR FECHA DE CIERRE")
    print("====================================")
    print(f"Fecha de cierre: {fecha_cierre.strftime('%d/%m/%Y')}")
    print(f"Filas antes del filtro: {filas_antes}")
    print(f"Filas eliminadas: {filas_eliminadas}")
    print(f"Ventas posteriores al cierre: {filas_posteriores}")
    print(f"Filas sin FechaOperacionComercial: {filas_sin_fecha}")
    print(f"Filas finales Base_Ajustada: {len(base_filtrada)}")

    return base_filtrada

# Genera el archivo de Base cobranzas ajustado
def generar_base_cobranzas_ajustada(base_consolidada):
    base_original = base_consolidada.copy()
    base_ajustada = base_consolidada.copy()

    base_ajustada = aplicar_fecha_cierre_base_ajustada(base_ajustada)
    base_ajustada = aplicar_ajuste_saldos_negativos(base_ajustada)

    fecha_cierre = obtener_fecha_cierre_ajuste()
    base_ajustada, _, _ = reclasificar_estado_por_fecha_cierre(base_ajustada, fecha_cierre)

    base_ajustada = aplicar_estado_parcial_cancelado(base_ajustada)
    base_ajustada = filtrar_ventas_hasta_fecha_cierre(base_ajustada)

    ruta_base_ajustada = SALIDA_LISTA / NOMBRE_BASE_AJUSTADA

    with pd.ExcelWriter(ruta_base_ajustada, engine="openpyxl") as writer:
        base_original.to_excel(
            writer,
            sheet_name=NOMBRE_HOJA_BASE_ORIGINAL_AJUSTADA,
            index=False,
        )

        base_ajustada.to_excel(
            writer,
            sheet_name=NOMBRE_HOJA_BASE_AJUSTADA,
            index=False,
        )

        formatear_hoja_base(writer.book[NOMBRE_HOJA_BASE_ORIGINAL_AJUSTADA])
        formatear_hoja_base(writer.book[NOMBRE_HOJA_BASE_AJUSTADA])

    print("\n✅ Base ajustada generada correctamente")

    return ruta_base_ajustada


# =========================================================
# 7. PROCESO PRINCIPAL
# =========================================================

def mostrar_rangos_descarga(rangos):
    print("\nRangos que se van a descargar:")

    for r in rangos:
        print(f"{r['anio']}: {r['fecha_inicio']} al {r['fecha_fin']}")


def ejecutar_descargas(rangos):
    if not DOWNLOADS_DIR.exists():
        raise FileNotFoundError(f"No existe la carpeta de Descargas: {DOWNLOADS_DIR}")

    driver = crear_driver_chrome()
    rangos_con_error = []

    try:
        if LOGIN_AUTOMATICO:
            iniciar_sesion_evolta(driver)
        else:
            driver.get(LOGIN_URL)
            input("\nInicia sesión en Evolta y luego presiona ENTER aquí... ")
            ir_a_reporte_cobranzas(driver)

        for rango in rangos:
            try:
                procesar_descarga_anual(driver, rango)
                time.sleep(2)

            except Exception as e:
                print(f"No se pudo descargar el año {rango['anio']}.")
                print("Motivo:")
                print(e)
                print("Se continuará con el siguiente año.")

                rangos_con_error.append({
                    "anio": rango["anio"],
                    "fecha_inicio": rango["fecha_inicio"],
                    "fecha_fin": rango["fecha_fin"],
                    "error": str(e),
                })

                try:
                    ir_a_reporte_cobranzas(driver)
                except Exception:
                    pass

                continue

    finally:
        driver.quit()

    return rangos_con_error


def mostrar_resumen_descargas(rangos_con_error):
    print("\n====================================")
    print("PROCESO DE DESCARGA TERMINADO")
    print("====================================")

    if rangos_con_error:
        print("\nRANGOS NO DESCARGADOS:")

        for error in rangos_con_error:
            print(
                f"{error['anio']}: "
                f"{error['fecha_inicio']} al {error['fecha_fin']} - "
                f"{error['error']}"
            )
    else:
        print("Todos los rangos fueron descargados correctamente.")

    print("\nArchivos guardados en:")
    print(ENTRADA_COBRANZAS.resolve())

# Ejecuta todo el proceso 
def main():
    print("====================================")
    print("AUTOMATIZACIÓN COBRANZAS EVOLTA")
    print("====================================")

    rangos = generar_rangos_por_anio(FECHA_INICIO_REPORTE, FECHA_FIN_REPORTE)
    mostrar_rangos_descarga(rangos)

    try:
        rangos_con_error = ejecutar_descargas(rangos)
        mostrar_resumen_descargas(rangos_con_error)

        ruta_base_consolidada = generar_base_cobranzas_consolidada()

        print("\nBase consolidada guardada en:")
        print(ruta_base_consolidada.resolve())

        base_consolidada = pd.read_excel(
            ruta_base_consolidada,
            sheet_name=NOMBRE_HOJA_BASE,
            engine="openpyxl",
        )

        base_consolidada = convertir_tipos_base(base_consolidada)
        ruta_base_ajustada = generar_base_cobranzas_ajustada(base_consolidada)

        print("\nBase ajustada guardada en:")
        print(ruta_base_ajustada.resolve())

    except Exception as e:
        print("\nOcurrió un error general:")
        print(e)
        raise


if __name__ == "__main__":
    main()
