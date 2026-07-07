import pandas as pd
from pathlib import Path
from bs4 import BeautifulSoup
import re

BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_LISTA_PRECIOS = BASE_DIR / "Flujo" / "Input"
SALIDA_LISTA = BASE_DIR / "Flujo" / "Output" / "Base_agentes_re_max.xlsx"

datos = []

def limpiar(texto):
    return re.sub(r"\s+", " ", texto).strip() if texto else ""

for archivo in ENTRADA_LISTA_PRECIOS.glob("*.html"):

    html = archivo.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")

    telefonos = soup.select("a[href^='tel:']")

    for tel in telefonos:
        telefono = tel.get("href", "").replace("tel:", "").strip()
        telefono = re.sub(r"\D", "", telefono)

        # subimos al contenedor del agente
        card = tel
        for _ in range(10):
            if card.parent:
                card = card.parent

            texto = limpiar(card.get_text(" "))

            # cuando el bloque ya tiene suficiente info, paramos
            if telefono in texto and len(texto) > 50:
                break

        texto = limpiar(card.get_text(" "))

        correo_match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", texto)
        correo = correo_match.group(0) if correo_match else ""

        # buscar nombre en etiquetas comunes
        nombre = ""
        for tag in card.find_all(["h1", "h2", "h3", "h4", "h5", "strong", "b", "span", "p"]):
            t = limpiar(tag.get_text())
            if (
                t
                and not re.search(r"\d|@|RE/MAX|REMAX|AGENTE|ASESOR|CONTACTAR|VER", t, re.I)
                and len(t.split()) >= 2
            ):
                nombre = t
                break

        grupo = ""
        grupo_match = re.search(r"(RE/MAX\s+[A-Za-zÁÉÍÓÚÑáéíóúñ0-9\s]+)", texto, re.I)
        if grupo_match:
            grupo = limpiar(grupo_match.group(1))

        datos.append({
            "Archivo": archivo.name,
            "Nombre": nombre,
            "Correo": correo,
            "Telefono": telefono,
            "Grupo": grupo,
            "Texto extraido": texto
        })

df = pd.DataFrame(datos).drop_duplicates(subset=["Telefono"])

SALIDA_LISTA.parent.mkdir(parents=True, exist_ok=True)
df.to_excel(SALIDA_LISTA, index=False)

print("Registros encontrados:", len(df))
print("Excel generado:", SALIDA_LISTA)
print(df.head())