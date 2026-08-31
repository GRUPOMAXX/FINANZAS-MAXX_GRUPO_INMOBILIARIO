import os
from dotenv import load_dotenv
import psycopg2
from datetime import datetime, timedelta

base_dir = os.getcwd()
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

# Fecha de cierre mensual
def obtener_ultimo_dia_mes_anterior():
    hoy = datetime.today()
    primer_dia_mes_actual = hoy.replace(day=1)
    ultimo_dia_mes_anterior = primer_dia_mes_actual - timedelta(days=1)
    return ultimo_dia_mes_anterior.strftime("%d/%m/%Y")


# =====================================================================================================
#                                       PARAMETROS A MODIFICAR
# =====================================================================================================


# ============================================ COBRANZAS ==============================================

# Proyectos Vigentes

PROYECTOS_VIGENTES_COBRANZAS = ["Prada", "Beyond", "Venecia"]    # Colocar los proyectos vigentes.
                                                                 #

# Inicio de Reporte

FECHA_INICIO_REPORTE = "01/01/2024"  # Fecha de Inicio del reporte. 

# Fecha de cierre mensual

FECHA_CORTE_REPORTE = "31/08/2026"   # Fecha de corte para el reporte.


PROYECTO = {
    "RESIDENCIAL PRADA": "Prada",
    "RESIDENCIAL BEYOND": "Beyond",
    "RESIDENCIAL VENECIA": "Venecia",
}



# ======================================================================================================
#                                             ¡NO TOCAR!
# ======================================================================================================


# Función fin de reporte
def sumar_anios_fecha(fecha_texto, anios):
    fecha = datetime.strptime(fecha_texto, "%d/%m/%Y")
    fecha_final = fecha.replace(year=fecha.year + anios)
    return fecha_final.strftime("%d/%m/%Y")


# Fecha fin
FECHA_FIN_REPORTE = sumar_anios_fecha(FECHA_CORTE_REPORTE, 5)


# Función año del reporte
def obtener_anio_mensual_reporte(fecha_corte_texto):
    fecha_corte = datetime.strptime(fecha_corte_texto, "%d/%m/%Y")
    return fecha_corte.year

# Año del reporte
ANIO_MENSUAL_REPORTE = obtener_anio_mensual_reporte(FECHA_CORTE_REPORTE)   # Año del reporte mensualizado.


# Tipo de Cambio SUNAT
RESULTADO_TIPO_CAMBIO =  os.path.join(base_dir, "Flujo", "output", f"tipo_cambio_sunat_{timestamp}.csv")
