"""
Web scraper de libros.

Recorre las páginas del catálogo de https://books.toscrape.com (un sitio
hecho para practicar scraping), extrae los datos de cada libro y los guarda
en Excel y CSV, con una hoja de resumen.

Uso:
    python3 scraper_libros.py                 # todas las páginas (50)
    python3 scraper_libros.py --paginas 3     # solo las primeras 3 páginas
"""

import argparse
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

URL_INICIAL = "https://books.toscrape.com/"
PAUSA_SEGUNDOS = 1  # Espera entre páginas para no sobrecargar el servidor
CALIFICACIONES = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


# ------------------------------------------------------------ 1. Descargar
def descargar(sesion, url, intentos=3):
    """Descarga una página. Si falla, reintenta unas veces antes de rendirse."""
    for intento in range(1, intentos + 1):
        try:
            respuesta = sesion.get(url, timeout=15)
            respuesta.raise_for_status()      # Error si el código no es 200
            respuesta.encoding = "utf-8"      # Para que "£" se lea bien
            return respuesta.text
        except requests.RequestException as error:
            print(f"  Intento {intento} falló ({error}).")
            if intento == intentos:
                raise
            time.sleep(2 * intento)


# ---------------------------------------------------------- 2. Extraer
def extraer_libros(html, url_pagina):
    """Saca los datos de cada libro de una página del catálogo."""
    soup = BeautifulSoup(html, "html.parser")
    libros = []

    for articulo in soup.select("article.product_pod"):
        enlace = articulo.select_one("h3 a")
        precio_texto = articulo.select_one("p.price_color").get_text(strip=True)
        estrellas = articulo.select_one("p.star-rating")["class"][1]  # p.ej. "Three"
        disponibilidad = articulo.select_one("p.availability").get_text(strip=True)

        libros.append({
            "titulo": enlace["title"],                         # el título completo
            "precio_gbp": float(precio_texto.replace("£", "")),
            "calificacion": CALIFICACIONES.get(estrellas),
            "disponible": disponibilidad == "In stock",
            "url": urljoin(url_pagina, enlace["href"]),
        })

    # ¿Hay una página siguiente? Si sí, devolvemos su dirección completa
    siguiente = soup.select_one("li.next a")
    url_siguiente = urljoin(url_pagina, siguiente["href"]) if siguiente else None
    return libros, url_siguiente


# ------------------------------------------------------------ 3. Recorrer
def recorrer_catalogo(url_inicial, max_paginas=None):
    sesion = requests.Session()
    sesion.headers["User-Agent"] = "Mozilla/5.0 (portfolio scraper; contact via GitHub)"

    todos, url, pagina = [], url_inicial, 0
    while url and (max_paginas is None or pagina < max_paginas):
        pagina += 1
        print(f"Página {pagina}: {url}")
        html = descargar(sesion, url)
        libros, url = extraer_libros(html, url)
        todos.extend(libros)
        print(f"  {len(libros)} libros (total: {len(todos)})")
        if url:
            time.sleep(PAUSA_SEGUNDOS)
    return pd.DataFrame(todos)


# ------------------------------------------------------------ 4. Exportar
def exportar(df, nombre_base):
    df = df.sort_values("titulo").reset_index(drop=True)
    df.to_csv(f"{nombre_base}.csv", index=False, encoding="utf-8-sig")

    resumen = pd.DataFrame({
        "Indicador": ["Libros extraídos", "Precio promedio (£)", "Precio mínimo (£)",
                      "Precio máximo (£)", "Calificación promedio", "Libros disponibles"],
        "Valor": [len(df), round(df["precio_gbp"].mean(), 2), df["precio_gbp"].min(),
                  df["precio_gbp"].max(), round(df["calificacion"].mean(), 2),
                  int(df["disponible"].sum())],
    })
    por_estrellas = (df.groupby("calificacion")
                       .agg(libros=("titulo", "count"), precio_promedio=("precio_gbp", "mean"))
                       .round(2).reset_index())
    mas_caros = df.nlargest(10, "precio_gbp")[["titulo", "precio_gbp", "calificacion"]]
    mejor_valorados = (df[df["calificacion"] == 5]
                       .nsmallest(10, "precio_gbp")[["titulo", "precio_gbp"]])

    archivo = f"{nombre_base}.xlsx"
    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        por_estrellas.to_excel(writer, sheet_name="Por calificación", index=False)
        mas_caros.to_excel(writer, sheet_name="Top 10 más caros", index=False)
        mejor_valorados.to_excel(writer, sheet_name="5 estrellas más baratos", index=False)
        df.to_excel(writer, sheet_name="Todos los libros", index=False)
    dar_formato(archivo)
    return archivo


def dar_formato(archivo):
    wb = load_workbook(archivo)
    for ws in wb.worksheets:
        for celda in ws[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="0F4C5C")
            celda.alignment = Alignment(horizontal="center")
            celda.value = str(celda.value).replace("_", " ").capitalize()
        for i, col in enumerate(ws.columns, start=1):
            largo = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(largo + 3, 60)
        ws.freeze_panes = "A2"
    wb.save(archivo)


# ------------------------------------------------------------------ Main
def main():
    parser = argparse.ArgumentParser(description="Scraper de books.toscrape.com")
    parser.add_argument("--paginas", type=int, default=None,
                        help="Número máximo de páginas a recorrer (por defecto, todas)")
    parser.add_argument("--url", default=URL_INICIAL, help="Página inicial del catálogo")
    parser.add_argument("--salida", default="libros", help="Nombre de los archivos de salida")
    args = parser.parse_args()

    df = recorrer_catalogo(args.url, args.paginas)
    if df.empty:
        print("No se encontraron libros. Revisa la URL o tu conexión.")
        return
    archivo = exportar(df, args.salida)
    print(f"\nListo: {len(df)} libros guardados en {archivo} y {args.salida}.csv")


if __name__ == "__main__":
    main()
