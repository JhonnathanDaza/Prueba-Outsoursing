"""PASO 5 - Extraer el texto de las facturas descargadas, armar un DataFrame por archivo y exportarlo a CSV.

Procedimiento (por cada PDF de descargas/):
  1. Abrir el PDF. La DIAN lo protege con contraseña = NIT del emisor o receptor consultado.
  2. Obtener el texto:
       - "texto": capa de texto del PDF con pdfplumber (exacta y rápida).
       - "ocr":   si el PDF no tiene capa de texto (es una imagen escaneada) o si se pide con
                  EXTRACCION=ocr, cada página se convierte en imagen (300 dpi) y se lee con
                  Tesseract OCR (palabras con su posición).
  3. Extraer las variables:
       - Datos del Documento / Datos del Emisor: numero_factura, fecha_emision, nit_emisor.
       - Tabla "Detalles de Productos": codigo, descripcion, cantidad, precio_unitario.
  4. Armar un DataFrame de pandas por archivo (una fila por producto) y exportarlo a
     salida/csv_por_factura/<numero_factura>.csv. También se arma el consolidado.
  5. Medir tiempos y calidad por archivo (salida/metricas_extraccion.csv).

Se puede ejecutar solo (sin volver a descargar):
    python paso5_extraccion.py            # capa de texto, OCR si hace falta
    python paso5_extraccion.py --ocr      # fuerza OCR sobre las imágenes de las páginas
    python paso5_extraccion.py --comparar # mide la exactitud del OCR contra la capa de texto
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from datetime import datetime
from pathlib import Path

from config import DESCARGAS, NIT_RESPALDO, SALIDA, log

CAMPOS_CSV = ["numero_factura", "fecha_emision", "nit_emisor", "codigo", "descripcion",
              "cantidad", "precio_unitario"]
EXTRACCION = os.getenv("EXTRACCION", "auto").lower()   # auto | texto | ocr
TESSERACT_RUTA = os.getenv("TESSERACT_RUTA", "")       # vacío = se busca solo
OCR_DPI = int(os.getenv("OCR_DPI", "300"))


# =====================================================================
# Librerías (se instalan solas la primera vez, como SeleniumBase)
# =====================================================================
def importar(modulo, paquete=None):
    try:
        return __import__(modulo)
    except ModuleNotFoundError:
        log(f"Paso 5 | Instalando {paquete or modulo} en {sys.executable} ...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", paquete or modulo],
                       capture_output=True)
        import importlib
        import site
        importlib.invalidate_caches()
        sys.path.append(site.getusersitepackages())
        return __import__(modulo)


pd = importar("pandas")
pdfplumber = importar("pdfplumber")


def sin_tildes(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


# =====================================================================
# Conversión de valores (formato colombiano: 1.234.567,89)
# =====================================================================
def a_numero(valor):
    """'$ 2.900.000,00' -> 2900000.0 ; '1,00' -> 1.0"""
    s = re.sub(r"[^\d,.]", "", str(valor or ""))
    if not s:
        return None
    if "," in s:  # la coma es el separador decimal
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def a_fecha(valor):
    """'05/11/2024' -> '2024-11-05' (ISO)."""
    for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(valor.strip(), formato).strftime("%Y-%m-%d")
        except (ValueError, AttributeError):
            pass
    return valor


def encabezado(texto):
    """Número de factura, fecha de emisión y NIT del emisor (tolerante a errores de OCR)."""
    plano = sin_tildes(texto)

    def buscar(patron):
        m = re.search(patron, plano, re.IGNORECASE)
        return m.group(1).strip() if m else None

    numero = buscar(r"numero\s+de\s+factura\s*[:;]?\s*([a-z0-9][a-z0-9\-]*)")
    fecha = buscar(r"fecha\s+de\s+emision\s*[:;]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})")
    nit = buscar(r"nit\s+del\s+emisor\s*[:;]?\s*([\d\.\- ]{6,20})")
    return {
        "numero_factura": numero.upper() if numero else None,
        "fecha_emision": a_fecha(fecha) if fecha else None,
        "nit_emisor": re.sub(r"\D", "", nit)[:15] if nit else None,
    }


# =====================================================================
# Método 1: capa de texto del PDF (pdfplumber)
# =====================================================================
def texto_celda(pagina, bbox):
    """Texto exacto de una celda. El PDF corta las descripciones a ancho fijo y conserva el
    espacio al final de la línea cuando el corte es entre palabras, así que basta con unir
    los caracteres de cada línea tal cual (sin agregar espacios)."""
    if not bbox:
        return ""
    lineas = {}
    for ch in pagina.crop(bbox).chars:
        lineas.setdefault(round(ch["top"]), []).append(ch)
    partes = ["".join(c["text"] for c in sorted(lineas[y], key=lambda c: c["x0"]))
              for y in sorted(lineas)]
    return re.sub(r"\s+", " ", "".join(partes)).strip()


def columnas_de(encabezados):
    """Índices de las columnas que interesan según los títulos de la tabla."""
    titulos = [sin_tildes(t or "").replace("\n", " ") for t in encabezados]

    def indice(*palabras, evitar=None):
        for i, t in enumerate(titulos):
            if all(p in t for p in palabras) and not (evitar and evitar in t):
                return i
        return None

    return {"nro": indice("nro"), "codigo": indice("codigo"), "descripcion": indice("descrip"),
            "cantidad": indice("cantidad"), "precio_unitario": indice("precio", "unitario", evitar="venta")}


def productos_texto(pdf):
    productos = []
    for pagina in pdf.pages:
        for tabla in pagina.find_tables():
            filas = tabla.extract()
            i_enc = next((i for i, f in enumerate(filas)
                          if any("cantidad" in sin_tildes(c or "") for c in f)
                          and any("codigo" in sin_tildes(c or "") for c in f)), None)
            if i_enc is None:
                continue  # no es la tabla de productos (p. ej. la de descuentos)
            col = columnas_de(filas[i_enc])
            for fila_obj, fila in zip(tabla.rows[i_enc + 1:], filas[i_enc + 1:]):
                if col["nro"] is None or not re.fullmatch(r"\d+", (fila[col["nro"]] or "").strip()):
                    continue
                celda = lambda k: texto_celda(pagina, fila_obj.cells[col[k]]) if col[k] is not None else ""
                productos.append({
                    "codigo": celda("codigo"),
                    "descripcion": celda("descripcion"),
                    "cantidad": a_numero(celda("cantidad")),
                    "precio_unitario": a_numero(celda("precio_unitario")),
                })
    return productos


# =====================================================================
# Método 2: OCR sobre la imagen de cada página (Tesseract)
# =====================================================================
def buscar_tesseract():
    candidatos = [TESSERACT_RUTA, shutil.which("tesseract") or "",
                  r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                  os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")]
    return next((c for c in candidatos if c and Path(c).exists()), None)


def limpiar_imagen(imagen):
    """Deja solo el texto oscuro: las líneas y fondos verdes de las tablas de la DIAN confunden
    al OCR (se pierde la fila de títulos y la primera fila de productos)."""
    import numpy as np
    from PIL import Image
    pixeles = np.asarray(imagen.convert("RGB"), dtype=np.int16)
    texto = pixeles.max(axis=2) < 150  # negro o gris oscuro; el verde claro y el blanco quedan fuera
    return Image.fromarray(np.where(texto, 0, 255).astype("uint8"))


def palabras_ocr(imagen, tesseract):
    """Palabras con su posición (TSV de Tesseract)."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta = Path(tmp) / "pagina.png"
        limpiar_imagen(imagen).save(ruta)
        idiomas = subprocess.run([tesseract, "--list-langs"], capture_output=True, text=True).stdout
        idioma = "spa" if "spa" in idiomas.split() else "eng"
        tsv = subprocess.run([tesseract, str(ruta), "stdout", "-l", idioma, "--psm", "6",
                              "-c", "tessedit_create_tsv=1"],
                             capture_output=True, text=True, encoding="utf-8").stdout
    palabras = []
    for linea in tsv.splitlines()[1:]:
        c = linea.split("\t")
        if len(c) == 12 and c[11].strip() and float(c[10]) > 0:
            palabras.append({"texto": c[11], "x": int(c[6]), "y": int(c[7]), "ancho": int(c[8]),
                             "alto": int(c[9]), "linea": (int(c[2]), int(c[3]), int(c[4]))})
    return palabras


def lineas_ocr(palabras):
    """Agrupa las palabras por renglón (misma altura)."""
    lineas = []
    for p in sorted(palabras, key=lambda p: (p["y"], p["x"])):
        if lineas and abs(lineas[-1]["y"] - p["y"]) <= p["alto"] * 0.6:
            lineas[-1]["palabras"].append(p)
        else:
            lineas.append({"y": p["y"], "palabras": [p]})
    for l in lineas:
        l["palabras"].sort(key=lambda p: p["x"])
        l["texto"] = " ".join(p["texto"] for p in l["palabras"])
    return lineas


def lineas_verticales(pixeles, y0, y1):
    """Posición x de las líneas verticales de la tabla (verde de la DIAN) entre y0 y y1."""
    import numpy as np
    franja = pixeles[max(int(y0), 0):int(y1)]
    if franja.shape[0] < 10:
        return []
    color = (franja[:, :, 0] < 150) & (franja[:, :, 1] > 170)
    xs = np.where(color.sum(axis=0) >= 0.5 * franja.shape[0])[0]
    grupos = []
    for x in xs:
        if grupos and x - grupos[-1][-1] <= 3:
            grupos[-1].append(x)
        else:
            grupos.append([x])
    return [sum(g) / len(g) for g in grupos]


def productos_ocr(palabras, pixeles=None):
    """Tabla de productos desde las palabras del OCR:
    - las columnas salen de la posición x de los títulos (Nro., Código, Descripción, ...);
    - cada fila se ancla en su número (columna Nro.): la descripción ocupa varias líneas
      centradas en la fila, así que cada palabra se asigna a la fila más cercana en y."""
    lineas = lineas_ocr(palabras)
    i_enc = next((i for i, l in enumerate(lineas)
                  if "cantidad" in sin_tildes(l["texto"]) and "codigo" in sin_tildes(l["texto"])), None)
    if i_enc is None:
        return []
    # Títulos: la fila detectada + las palabras a pocos píxeles (algunos títulos quedan
    # un poco más arriba o más abajo, como "Descripción", y el OCR los pone en otra línea)
    y_enc = lineas[i_enc]["y"]
    enc = lineas[i_enc]["palabras"] + [p for p in palabras if p not in lineas[i_enc]["palabras"]
                                       and abs(p["y"] - y_enc) <= 25]
    x_de = lambda palabra: next((p["x"] for p in enc if palabra in sin_tildes(p["texto"])), None)
    inicio = {"nro": x_de("nro"), "codigo": x_de("codigo"), "descripcion": x_de("descrip"),
              "um": x_de("u/m"), "cantidad": x_de("cantidad"), "precio": x_de("precio"),
              "fin": x_de("descuento")}
    if None in (inicio["nro"], inicio["codigo"], inicio["descripcion"], inicio["cantidad"], inicio["precio"]):
        return []
    inicio = {k: v for k, v in inicio.items() if v is not None}
    # Límite izquierdo de cada columna: a mitad de camino entre títulos (los valores van centrados)
    orden = sorted(inicio.items(), key=lambda kv: kv[1])
    cortes = [(orden[0][1] - 10 ** 6, orden[0][0])]
    for (k_ant, x_ant), (k, x) in zip(orden, orden[1:]):
        ancho_titulo = next((p["ancho"] for p in enc if p["x"] == x_ant), 0)
        cortes.append(((x_ant + ancho_titulo + x) / 2 if k_ant == "nro" else x - 15, k))

    def columna_por_titulos(p):
        centro = p["x"] + p["ancho"] / 2 if p["ancho"] < 60 else p["x"]
        nombre = cortes[0][1]
        for limite, k in cortes:
            if centro >= limite:
                nombre = k
        return nombre

    columna_lineas = None

    y_ini = max(p["y"] + p["alto"] for p in enc)
    y_fin = next((l["y"] for l in lineas[i_enc + 1:]
                  if re.search(r"notas finales|referencias|descuentos y recargos|^nro\.? tipo|subtotal|total|hoja \d",
                               sin_tildes(l["texto"]))), 10 ** 9)
    tabla = [p for p in palabras if y_ini < p["y"] < y_fin]

    # Columnas exactas con las líneas verticales de la tabla (si se ven en la imagen):
    # cada palabra va a la celda entre dos líneas; la celda toma el nombre de su título.
    if pixeles is not None and tabla:
        fondo = min(y_fin, max(p["y"] + p["alto"] for p in tabla) + 20)
        lineas_v = lineas_verticales(pixeles, y_enc - 5, fondo)
        if len(lineas_v) >= 6:
            nombres = {}
            for i, (x0, x1) in enumerate(zip(lineas_v, lineas_v[1:])):
                for k, x in inicio.items():
                    ancho = next((p["ancho"] for p in enc if p["x"] == x), 0)
                    if x0 <= x + ancho / 2 < x1 and i not in nombres:
                        nombres[i] = k

            def columna_lineas(p):
                centro = p["x"] + p["ancho"] / 2
                for i, (x0, x1) in enumerate(zip(lineas_v, lineas_v[1:])):
                    if x0 <= centro < x1:
                        return nombres.get(i, "fin")
                return "fin"

    columna = columna_lineas or columna_por_titulos
    centro_y = lambda p: p["y"] + p["alto"] / 2
    anclas = sorted((centro_y(p), p["texto"]) for p in tabla
                    if columna(p) == "nro" and re.fullmatch(r"\d{1,3}", p["texto"]))
    if not anclas:
        return []
    filas = [{"y": y, "celdas": {}} for y, _ in anclas]
    for p in sorted(tabla, key=lambda p: (p["y"], p["x"])):
        col = columna(p)
        if col in ("nro", "um", "fin"):
            continue
        fila = min(filas, key=lambda f: abs(f["y"] - centro_y(p)))
        fila["celdas"].setdefault(col, []).append(p)

    def unir(celda):
        """Texto de una celda en orden de lectura: renglón por renglón, de izquierda a derecha."""
        renglones = []
        for p in sorted(celda, key=lambda p: centro_y(p)):
            if renglones and abs(centro_y(renglones[-1][0]) - centro_y(p)) <= p["alto"] * 0.5:
                renglones[-1].append(p)
            else:
                renglones.append([p])
        return " ".join(p["texto"] for r in renglones for p in sorted(r, key=lambda p: p["x"])).strip()

    return [{"codigo": unir(f["celdas"].get("codigo", [])),
             "descripcion": unir(f["celdas"].get("descripcion", [])),
             "cantidad": a_numero(unir(f["celdas"].get("cantidad", []))),
             "precio_unitario": a_numero(unir(f["celdas"].get("precio", [])))} for f in filas]


def leer_con_ocr(pdf, tesseract):
    texto, productos = [], []
    for pagina in pdf.pages:
        imagen = pagina.to_image(resolution=OCR_DPI).original
        import numpy as np
        palabras = palabras_ocr(imagen, tesseract)
        texto.append("\n".join(l["texto"] for l in lineas_ocr(palabras)))
        productos += productos_ocr(palabras, np.asarray(imagen.convert("RGB"), dtype=np.int16))
    return "\n".join(texto), productos


# =====================================================================
# Una factura -> un DataFrame
# =====================================================================
def abrir_pdf(ruta, claves):
    ultimo = None
    for clave in dict.fromkeys(c for c in claves if c is not None):
        try:
            return pdfplumber.open(ruta, password=clave)
        except Exception as e:  # contraseña incorrecta
            ultimo = e
    raise ValueError(f"No pude abrir el PDF con las contraseñas probadas ({type(ultimo).__name__})")


def extraer_factura(ruta, claves=("",), modo=EXTRACCION):
    """Devuelve (DataFrame, métricas) de una factura."""
    t0 = time.time()
    with abrir_pdf(ruta, claves) as pdf:
        paginas = len(pdf.pages)
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
        tiene_texto = len(texto.strip()) > 50
        t_lectura = time.time() - t0
        if modo == "ocr" or (modo == "auto" and not tiene_texto):
            tesseract = buscar_tesseract()
            if not tesseract:
                raise RuntimeError("El PDF no tiene capa de texto y no está instalado Tesseract OCR "
                                   "(define TESSERACT_RUTA en .env)")
            metodo = "ocr"
            texto, productos = leer_con_ocr(pdf, tesseract)
        else:
            metodo = "texto"
            productos = productos_texto(pdf)
    datos = encabezado(texto)
    df = pd.DataFrame([{**datos, **p} for p in productos] or [datos], columns=CAMPOS_CSV)
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")
    df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce")

    enc_ok = sum(v is not None for v in datos.values())
    # El código puede venir vacío en la factura original: no cuenta como error, se reporta aparte
    prod_ok = sum(bool(p["descripcion"]) and p["cantidad"] is not None
                  and p["precio_unitario"] is not None for p in productos)
    sin_codigo = sum(not p["codigo"] for p in productos)
    completo = enc_ok == 3 and productos and prod_ok == len(productos)
    metricas = {
        "pdf": ruta.name, "numero_factura": datos["numero_factura"], "metodo": metodo,
        "paginas": paginas, "productos": len(productos), "campos_encabezado": f"{enc_ok}/3",
        "productos_completos": f"{prod_ok}/{len(productos)}", "productos_sin_codigo": sin_codigo,
        "estado": "OK" if completo else "REVISAR",
        "t_lectura_s": round(t_lectura, 3), "t_extraccion_s": round(time.time() - t0, 3),
    }
    return df, metricas


def paso5_extraer_todo(nit="", driver=None, modo=EXTRACCION):
    """Una tabla (DataFrame) por PDF, exportada a su propio CSV, más el consolidado.
    `driver` ya no se usa (la extracción no necesita el navegador); se deja por compatibilidad."""
    pdfs = sorted(DESCARGAS.glob("*.pdf"))
    if not pdfs:
        log("Paso 5 | No hay PDFs en descargas/")
        return []
    carpeta_csv = SALIDA / "csv_por_factura"
    carpeta_csv.mkdir(parents=True, exist_ok=True)
    tablas, metricas = [], []
    for ruta in pdfs:
        try:
            df, m = extraer_factura(ruta, [nit, NIT_RESPALDO, ""], modo)
            nombre = re.sub(r"[^\w\-]", "_", str(m["numero_factura"] or ruta.stem[:16]))
            df.to_csv(carpeta_csv / f"{nombre}.csv", index=False, encoding="utf-8-sig")
            tablas.append(df.assign(archivo_pdf=ruta.name))
            m["csv"] = f"{nombre}.csv"
            log(f"Paso 5 | {ruta.name[:16]}... -> {nombre}.csv ({m['metodo']}, {m['estado']}, "
                f"{m['productos']} productos, {m['t_extraccion_s']} s)")
        except Exception as e:
            m = {"pdf": ruta.name, "estado": "ERROR", "error": f"{type(e).__name__}: {e}"[:200]}
            log(f"Paso 5 | Error en {ruta.name[:16]}...: {m['error']}")
        metricas.append(m)
    if tablas:
        pd.concat(tablas, ignore_index=True).to_csv(SALIDA / "facturas_consolidado.csv",
                                                    index=False, encoding="utf-8-sig")
    return metricas


def comparar_metodos(nit=NIT_RESPALDO):
    """Extrae cada PDF con los dos métodos y mide la exactitud del OCR campo por campo,
    tomando la capa de texto como referencia. Guarda salida/comparacion_texto_vs_ocr.csv."""
    import difflib
    norm = lambda v: re.sub(r"\s+", "", "" if pd.isna(v) else str(v)).upper()
    filas, tiempos = [], {"texto": 0.0, "ocr": 0.0}
    for ruta in sorted(DESCARGAS.glob("*.pdf")):
        df_t, m_t = extraer_factura(ruta, [nit, ""], "texto")
        df_o, m_o = extraer_factura(ruta, [nit, ""], "ocr")
        tiempos["texto"] += m_t["t_extraccion_s"]
        tiempos["ocr"] += m_o["t_extraccion_s"]
        for i in range(max(len(df_t), len(df_o))):
            t = df_t.iloc[i] if i < len(df_t) else pd.Series(dtype=object)
            o = df_o.iloc[i] if i < len(df_o) else pd.Series(dtype=object)
            fila = {"pdf": ruta.name[:16], "item": i + 1}
            for c in CAMPOS_CSV:
                fila[c] = norm(t.get(c)) == norm(o.get(c))
            fila["similitud_descripcion"] = round(difflib.SequenceMatcher(
                None, norm(t.get("descripcion")), norm(o.get("descripcion"))).ratio(), 3)
            filas.append(fila)
    comp = pd.DataFrame(filas)
    SALIDA.mkdir(exist_ok=True)
    comp.to_csv(SALIDA / "comparacion_texto_vs_ocr.csv", index=False, encoding="utf-8-sig")
    resumen = {c: round(100 * comp[c].mean(), 1) for c in CAMPOS_CSV}
    resumen["descripcion_similitud"] = round(100 * comp["similitud_descripcion"].mean(), 1)
    print("\nEXACTITUD DEL OCR (referencia: capa de texto del PDF)")
    for k, v in resumen.items():
        print(f"  {k:24} {v:5.1f} %")
    n = len(list(DESCARGAS.glob("*.pdf")))
    print(f"\nTiempo capa de texto: {tiempos['texto']:.2f} s ({tiempos['texto'] / max(n, 1):.2f} s/PDF)")
    print(f"Tiempo OCR:           {tiempos['ocr']:.2f} s ({tiempos['ocr'] / max(n, 1):.2f} s/PDF)")
    print(f"Detalle en {SALIDA / 'comparacion_texto_vs_ocr.csv'}")
    return resumen


if __name__ == "__main__":
    # Solo extracción (sin descargar): python paso5_extraccion.py [--ocr]
    if "--comparar" in sys.argv:
        comparar_metodos()
        sys.exit(0)
    inicio = time.time()
    modo = "ocr" if "--ocr" in sys.argv else EXTRACCION
    met = paso5_extraer_todo(NIT_RESPALDO, modo=modo)
    ok = sum(m.get("estado") == "OK" for m in met)
    print(f"\nPDFs: {len(met)} | OK: {ok} | REVISAR/ERROR: {len(met) - ok} | "
          f"Tiempo total: {time.time() - inicio:.2f} s | CSV en {SALIDA / 'csv_por_factura'}")
    if met:
        from config import guardar_csv
        guardar_csv(SALIDA / "metricas_extraccion.csv", met, list(dict.fromkeys(k for m in met for k in m)))
