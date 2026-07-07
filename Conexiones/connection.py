import os
from dotenv import load_dotenv
import psycopg2
from datetime import datetime

base_dir = os.getcwd()
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

#MODELO DE FINANZAS
# Parámetros a modificar

finanzas_anlo_modelo = '2025'  
finanzas_mes_modelo = '08'  


# Tipo de Cambio SUNAT
RESULTADO_TIPO_CAMBIO =  os.path.join(base_dir, "Flujo", "output", f"tipo_cambio_sunat_{timestamp}.csv")
