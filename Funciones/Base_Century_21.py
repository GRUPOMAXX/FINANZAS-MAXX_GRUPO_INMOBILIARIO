from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from pathlib import Path

# ==============================
# RUTA DE SALIDA
# ==============================
BASE_DIR = Path(__file__).resolve().parent.parent
SALIDA_LISTA = BASE_DIR / "Flujo" / "Output" / "Base_Century_21.xlsx"
SALIDA_LISTA.parent.mkdir(parents=True, exist_ok=True)

# ==============================
# INICIAR NAVEGADOR
# ==============================
driver = webdriver.Chrome()

# ==============================
# VARIABLES
# ==============================
agentes = []
pagina = 1

# ==============================
# SCRAPING
# ==============================
while True:
    url = f"https://century21.pe/asesores?estado=Lima&pagina={pagina}"
    print(f" 🔎 Procesando página {pagina}")

    driver.get(url)
    time.sleep(3)

    cards = driver.find_elements(By.CSS_SELECTOR, "div[class*='col']")

    nuevos_en_pagina = 0  # 🔴 contador real

    for card in cards:
        texto = [t.strip() for t in card.text.split("\n") if t.strip()]

        # 🔴 validar que sea un agente real (no contenedor vacío)
        tiene_datos = any("@" in t or "+" in t for t in texto)

        if len(texto) >= 2 and tiene_datos:
            nombre = texto[0]
            oficina = texto[1]

            telefono = ""
            email = ""

            for t in texto:
                if "+" in t:
                    telefono = t
                if "@" in t:
                    email = t

            agentes.append([nombre, oficina, telefono, email])
            nuevos_en_pagina += 1

    # ==============================
    # CONDICIÓN DE PARADA REAL
    # ==============================
    if nuevos_en_pagina == 0:
        print(f" 🚫 Página {pagina} sin datos. Fin del scraping.")
        break

    pagina += 1

# ==============================
# CERRAR NAVEGADOR
# ==============================
driver.quit()

# ==============================
# CREAR BASE DE DATOS
# ==============================
df = pd.DataFrame(agentes, columns=["Nombre", "Oficina", "Teléfono", "Email"])

# eliminar duplicados
df.drop_duplicates(inplace=True)

# guardar archivo
df.to_excel(SALIDA_LISTA, index=False)

print("===================================")
print(" ✅ Base completa generada")
print(f" 📁 Total registros: {len(df)}")
print(f"Archivo: {SALIDA_LISTA}")
print("===================================")