import requests
import pandas as pd
from datetime import datetime
import os
import time
import sys

# Ruta completa donde estás ejecutando el script
base_dir = os.getcwd() 

# Agregar la carpeta padre al path para poder importar módulos desde allí
base_scr = sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

anio_inicio = 2023  #Año de inicio

from Conexiones.connection import RESULTADO_TIPO_CAMBIO


URL = "https://e-consulta.sunat.gob.pe/cl-at-ittipcam/tcS01Alias/listarTipoCambio"
TOKEN = "yw3jiixf6cn49sxtznj30cos9zc1dgxhoxpldc5ynbr8o4ola1t9"
ARCHIVO = RESULTADO_TIPO_CAMBIO

def obtener_tipo_cambio(anio: int, mes: int, reintentos=3):
    """Consulta un mes específico de un año, con reintentos automáticos."""
    payload = {"anio": anio, "mes": mes, "token": TOKEN}
    headers = {
        "Content-Type": "application/json; charset=UTF-8",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Origin": "https://e-consulta.sunat.gob.pe",
        "Referer": "https://e-consulta.sunat.gob.pe/cl-at-ittipcam/tcS01Alias",
        "Cookie":"TS0105b0c6=014dc399cb03db2bc44d05e424decd57ab9a86a85eb5b5cd888466dab15d393e1e4d0e20fac3fc4f79a25dae3a45e74c26615f0eba; Path=/; Domain=e-consulta.sunat.gob.pe;",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    }

    for intento in range(1, reintentos + 1):
        try:
            res = requests.post(URL, json=payload, headers=headers, timeout=20)
            res.raise_for_status()
            data = res.json()
            print(f"  ✅ {anio}-{mes+1:02d}: {len(data)} registros")
            return data
        except requests.exceptions.Timeout:
            print(f"  ⚠️ Timeout {anio}-{mes+1:02d} (intento {intento}/{reintentos})")
        except Exception as e:
            print(f"  ⚠️ Error {anio}-{mes+1:02d}: {e}")

        # Espera 3 segundos antes de volver a intentar
        time.sleep(3)

    print(f"  ❌ Falló {anio}-{mes+1:02d} después de {reintentos} intentos")
    return []

def construir_base_segura():
    hoy = datetime.now().date()

    # Si ya existe un CSV previo, lo cargamos
    if os.path.exists(ARCHIVO):
        df = pd.read_csv(ARCHIVO, parse_dates=["Fecha"])
        df["Fecha"] = df["Fecha"].dt.date
    else:
        df = pd.DataFrame(columns=["Fecha", "Compra", "Venta"])

    # Determinar desde qué año continuar
    if not df.empty:
        ultimo_anio = df["Fecha"].max().year
    else:
        ultimo_anio = anio_inicio  # desde este año en adelante

    for anio in range(ultimo_anio, hoy.year + 1):
        print(f"\n📅 Consultando año {anio}...")
        for mes in range(0, 12):
            # Si estamos en el año actual y mes futuro, no consultamos
            if anio == hoy.year and mes + 1 > hoy.month:
                continue

            datos = obtener_tipo_cambio(anio, mes)
            filas = {}

            for item in datos:
                fecha = datetime.strptime(item["fecPublica"], "%d/%m/%Y").date()
                if item["codTipo"] == "C":
                    filas.setdefault(fecha, {})["Compra"] = float(item["valTipo"])
                elif item["codTipo"] == "V":
                    filas.setdefault(fecha, {})["Venta"] = float(item["valTipo"])

            nuevos = [
                {"Fecha": f, "Compra": v.get("Compra"), "Venta": v.get("Venta")}
                for f, v in filas.items()
            ]

            if nuevos:
                df_nuevo = pd.DataFrame(nuevos)
                df = pd.concat([df, df_nuevo]).drop_duplicates(subset=["Fecha"], keep="last")

            # Pausa corta entre meses (evita bloqueo)
            time.sleep(1.5)

        # Pausa un poco más larga entre años
        time.sleep(3)

    df = df.sort_values("Fecha")
    df.to_csv(ARCHIVO, index=False, encoding="utf-8-sig")
    print(f"\n✅ Base actualizada hasta {df['Fecha'].max()} ({len(df)} registros)")

if __name__ == "__main__":
    construir_base_segura()

