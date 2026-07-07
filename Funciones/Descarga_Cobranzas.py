# Programa para descargar los reportes de Cobranzas
# Creado por Eduardo Miguel Huamani Acosta              06/07/26

from pathlib import Path
from datetime import datetime, date
import time
import shutil
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Reporte_Cobranzas import FECHA_INICIO_REPORTE, FECHA_FIN_REPORTE

# =========================================================
# 1. VARIABLES QUE PUEDES MODIFICAR
# =========================================================

LOGIN_URL = "https://v4.evolta.pe/Login/Acceso/Index"
REPORTE_COBRANZA_URL = "https://v4.evolta.pe/Reportes/ReporteCobranza/Index"

# Carpeta donde Chrome descarga primero
DOWNLOADS_DIR = Path(r"C:\Users\ehuamani\Downloads")   # Modificar segun la ruta de descarga del usuario 

# CAMBIA ESTA RUTA POR LA RUTA REAL DE TU PROYECTO
BASE_DIR = Path(__file__).resolve().parent.parent

# Carpeta final donde quieres que queden los archivos
ENTRADA_COBRANZAS = BASE_DIR / "Flujo" / "Input" / "Cobranzas"
ENTRADA_COBRANZAS.mkdir(parents=True, exist_ok=True)

# Formato que se selecciona en Evolta
# Puede ser "Csv" o "Excel"
FORMATO_EVOLTA = "Csv"

# Tu CSV está delimitado por coma
SEPARADOR_CSV = ","

# Si descarga CSV, también lo convierte a Excel seguro
CONVERTIR_CSV_A_EXCEL = True

# Reemplaza archivos existentes
OVERWRITE_EXISTING = True

# Si quieres borrar el archivo de Descargas después de copiarlo, pon True
BORRAR_ORIGINAL_DOWNLOADS = False


# =========================================================
# 2. FUNCIONES DE FECHA
# =========================================================

def convertir_fecha(fecha_texto: str) -> date:
    return datetime.strptime(fecha_texto, "%d/%m/%Y").date()


def fecha_texto(fecha: date) -> str:
    return fecha.strftime("%d/%m/%Y")


def generar_rangos_por_anio(fecha_inicio_txt: str, fecha_fin_txt: str):
    inicio = convertir_fecha(fecha_inicio_txt)
    fin = convertir_fecha(fecha_fin_txt)

    if inicio > fin:
        raise ValueError("❌ La fecha de inicio no puede ser mayor que la fecha de fin.")

    rangos = []

    for anio in range(inicio.year, fin.year + 1):
        inicio_anio = date(anio, 1, 1)
        fin_anio = date(anio, 12, 31)

        desde = max(inicio, inicio_anio)
        hasta = min(fin, fin_anio)

        rangos.append({
            "anio": anio,
            "fecha_inicio": fecha_texto(desde),
            "fecha_fin": fecha_texto(hasta)
        })

    return rangos


# =========================================================
# 3. FUNCIONES PARA ENTRAR A EVOLTA
# =========================================================

def esperar_elemento(driver, xpath, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )


def seleccionar_formato(driver, formato: str):
    if formato not in ["Csv", "Excel"]:
        raise ValueError("El formato debe ser 'Csv' o 'Excel'.")

    print(f"Seleccionando formato en Evolta: {formato}")

    xpath_radio = f"//*[normalize-space()='{formato}']/preceding::input[@type='radio'][1]"

    radio = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, xpath_radio))
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", radio)
    time.sleep(0.5)


def escribir_fecha(driver, nombre_campo: str, fecha: str):
    print(f"Escribiendo {nombre_campo}: {fecha}")

    xpath = f"//*[normalize-space()='{nombre_campo}']/following::input[1]"
    campo = esperar_elemento(driver, xpath, timeout=20)

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

    time.sleep(0.8)


def click_exportar(driver):
    print("Buscando botón Exportar...")

    xpath_exportar = """
    //button[contains(normalize-space(.), 'Exportar')]
    | //a[contains(normalize-space(.), 'Exportar')]
    | //input[@value='Exportar']
    | //*[@title='Exportar']
    """

    boton = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, xpath_exportar))
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
    time.sleep(0.5)

    print("Haciendo clic en Exportar...")
    driver.execute_script("arguments[0].click();", boton)

    time.sleep(2)


# =========================================================
# 4. CONTROL DE DESCARGAS
# =========================================================

def configurar_descargas_chrome(driver):
    ruta_descarga = str(DOWNLOADS_DIR.resolve())

    try:
        driver.execute_cdp_cmd(
            "Browser.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": ruta_descarga
            }
        )
    except Exception:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": ruta_descarga
            }
        )



def obtener_archivos_downloads():
    extensiones = [".csv", ".xlsx", ".xls"]

    return {
        archivo.name
        for archivo in DOWNLOADS_DIR.iterdir()
        if archivo.is_file()
        and archivo.suffix.lower() in extensiones
    }


def esperar_nueva_descarga_downloads(archivos_antes, tiempo_inicio, timeout=180):
    print("⌛ Esperando archivo nuevo en Descargas...")

    tiempo_limite = time.time() + timeout

    while time.time() < tiempo_limite:
        temporales = list(DOWNLOADS_DIR.glob("*.crdownload"))

        if temporales:
            time.sleep(1)
            continue

        candidatos = []

        for archivo in DOWNLOADS_DIR.iterdir():
            if not archivo.is_file():
                continue

            if archivo.suffix.lower() not in [".csv", ".xlsx", ".xls"]:
                continue

            nombre_nuevo = archivo.name not in archivos_antes
            modificado_reciente = archivo.stat().st_mtime >= tiempo_inicio

            if nombre_nuevo or modificado_reciente:
                candidatos.append(archivo)

        if candidatos:
            archivo_nuevo = max(candidatos, key=lambda f: f.stat().st_mtime)
            return archivo_nuevo

        time.sleep(1)

    raise TimeoutError("❌ No se detectó ningún archivo nuevo en Descargas.")


def renombrar_en_downloads(archivo_descargado: Path, anio: int):
    extension = archivo_descargado.suffix.lower()

    if extension not in [".csv", ".xlsx", ".xls"]:
        raise ValueError(f"⚠️ Extensión no esperada: {extension}")

    nuevo_nombre_downloads = DOWNLOADS_DIR / f"Cobranzas_{anio}{extension}"

    if nuevo_nombre_downloads.exists() and OVERWRITE_EXISTING:
        nuevo_nombre_downloads.unlink()

    if nuevo_nombre_downloads.exists() and not OVERWRITE_EXISTING:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nuevo_nombre_downloads = DOWNLOADS_DIR / f"Cobranzas_{anio}_{timestamp}{extension}"

    archivo_descargado.rename(nuevo_nombre_downloads)

    if not nuevo_nombre_downloads.exists():
        raise FileNotFoundError(f"❌ No se renombró correctamente en Descargas: {nuevo_nombre_downloads}")

    return nuevo_nombre_downloads


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


# =========================================================
# 5. CONVERSIÓN CSV A EXCEL 
# =========================================================

def leer_csv_seguro(ruta_csv: Path) -> pd.DataFrame:
    codificaciones = ["utf-8-sig", "latin1", "cp1252"]

    for encoding in codificaciones:
        try:
            df = pd.read_csv(
                ruta_csv,
                sep=SEPARADOR_CSV,
                dtype=str,
                encoding=encoding,
                keep_default_na=False
            )
            return df
        except Exception:
            continue

    raise ValueError(f"No se pudo leer el CSV: {ruta_csv}")


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

    print(f"✅ Excel Generado: {ruta_csv.name}")
    print(ruta_excel)

    return ruta_excel


# =========================================================
# 6. LIMPIEZA DE ARCHIVOS PREVIOS
# =========================================================

def limpiar_archivos_previos_anio(anio: int):
    if not OVERWRITE_EXISTING:
        return

    posibles = [
        DOWNLOADS_DIR / f"Cobranzas_{anio}.csv",
        DOWNLOADS_DIR / f"Cobranzas_{anio}.xlsx",
        ENTRADA_COBRANZAS / f"Cobranzas_{anio}.csv",
        ENTRADA_COBRANZAS / f"Cobranzas_{anio}.xlsx",
    ]

    for archivo in posibles:
        if archivo.exists():
            try:
             archivo.unlink()
            except Exception:
                pass


# =========================================================
# 7. PROGRAMA PRINCIPAL
# =========================================================

def main():
    print("====================================")
    print("AUTOMATIZACIÓN COBRANZAS EVOLTA")
    print("====================================")

    if not DOWNLOADS_DIR.exists():
        raise FileNotFoundError(f"❌ No existe la carpeta de Descargas: {DOWNLOADS_DIR}")

    rangos = generar_rangos_por_anio(FECHA_INICIO_REPORTE, FECHA_FIN_REPORTE)

    print("\nRangos que se van a descargar:")
    for r in rangos:
        print(f"{r['anio']}: {r['fecha_inicio']} al {r['fecha_fin']}")

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": str(DOWNLOADS_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    }

    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    configurar_descargas_chrome(driver)

    try:
        driver.get(LOGIN_URL)

        input(
            "\nInicia sesión en Evolta, presiona ENTER aquí... "
        )

        driver.get(REPORTE_COBRANZA_URL)

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//button[contains(normalize-space(.), 'Exportar')]"
                    " | //a[contains(normalize-space(.), 'Exportar')]"
                    " | //input[@value='Exportar']"
                    " | //*[@title='Exportar']"
                )
            )
        )

        for r in rangos:
            anio = r["anio"]
            fecha_inicio = r["fecha_inicio"]
            fecha_fin = r["fecha_fin"]

            print("\n====================================")
            print(f"PROCESANDO AÑO {anio}")
            print(f"Fecha inicio: {fecha_inicio}")
            print(f"Fecha fin: {fecha_fin}")
            print("====================================")

            limpiar_archivos_previos_anio(anio)

            archivos_antes = obtener_archivos_downloads()

            seleccionar_formato(driver, FORMATO_EVOLTA)

            escribir_fecha(driver, "Fecha de inicio", fecha_inicio)
            escribir_fecha(driver, "Fecha de fin", fecha_fin)

            tiempo_inicio_descarga = time.time()

            click_exportar(driver)

            archivo_descargado_downloads = esperar_nueva_descarga_downloads(
                archivos_antes=archivos_antes,
                tiempo_inicio=tiempo_inicio_descarga,
                timeout=180
            )

            archivo_renombrado_downloads = renombrar_en_downloads(
                archivo_descargado=archivo_descargado_downloads,
                anio=anio
            )

            archivo_final = copiar_renombrado_a_cobranzas(
                archivo_renombrado=archivo_renombrado_downloads
            )

            if CONVERTIR_CSV_A_EXCEL and archivo_final.suffix.lower() == ".csv":
                convertir_csv_a_excel_seguro(archivo_final)

            time.sleep(2)

        print("\n====================================")
        print("PROCESO TERMINADO CORRECTAMENTE")
        print("====================================")

        print("\nArchivos renombrados en Descargas:")
        print(DOWNLOADS_DIR.resolve())
        print("\n")

    except Exception as e:
        print("\nOcurrió un error:")
        print(e)

    finally:
        driver.quit()


if __name__ == "__main__":
    main()