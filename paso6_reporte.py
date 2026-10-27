"""PASO 6 - Mostrar y guardar tiempos y métricas."""

from config import SALIDA, guardar_csv, log

def paso6_reporte(tiempos, met_descarga, met_extraccion):
    SALIDA.mkdir(exist_ok=True)
    if met_descarga:
        guardar_csv(SALIDA / "metricas_descarga.csv", met_descarga, list(met_descarga[0]))
    if met_extraccion:
        campos = list(dict.fromkeys(k for m in met_extraccion for k in m))
        guardar_csv(SALIDA / "metricas_extraccion.csv", met_extraccion, campos)

    exitosos = sum(m["estado"] == "Exitoso" for m in met_descarga)
    leidos = [m for m in met_extraccion if m.get("estado") != "ERROR"]
    t_ext = [m["t_extraccion_s"] for m in leidos]
    metodos = {k: sum(m.get("metodo") == k for m in leidos) for k in ("texto", "ocr")}
    completos = sum(m["estado"] == "OK" for m in met_extraccion)
    t_captcha = sum(m["t_captcha_s"] for m in met_descarga)
    intentos = sum(m["intentos"] for m in met_descarga)
    resumen = (
        "MÉTRICAS\n"
        f"Facturas consultadas: {len(met_descarga)}\n"
        f"Procesadas (Exitoso): {exitosos}\n"
        f"No procesadas: {len(met_descarga) - exitosos}\n"
        f"Tasa de éxito de descarga: {exitosos / max(len(met_descarga), 1):.0%}\n"
        f"Intentos totales: {intentos}\n"
        f"PDFs con datos completos: {completos} de {len(met_extraccion)}\n"
        f"Productos extraídos: {sum(m.get('productos', 0) for m in leidos)}\n"
        f"Método de extracción: capa de texto {metodos['texto']}, OCR {metodos['ocr']}\n"
        f"Productos sin código en la factura: {sum(m.get('productos_sin_codigo', 0) for m in leidos)}\n"
        "\nTIEMPOS\n"
        + "".join(f"Tiempo {etapa}: {seg:.1f} s\n" for etapa, seg in tiempos.items())
        + f"Tiempo en captcha: {t_captcha:.1f} s\n"
        + f"Promedio por factura: {tiempos.get('descarga', 0) / max(len(met_descarga), 1):.1f} s\n"
        + (f"Extracción por PDF: promedio {sum(t_ext) / len(t_ext):.2f} s, "
           f"mínimo {min(t_ext):.2f} s, máximo {max(t_ext):.2f} s\n" if t_ext else "")
        + "\nRESTRICCIONES\n" + "".join(f"- {r}\n" for r in RESTRICCIONES)
    )
    (SALIDA / "reporte.txt").write_text(resumen, encoding="utf-8")
    log("Paso 6 | RESUMEN\n" + resumen)
    return resumen


RESTRICCIONES = [
    "El sitio de la DIAN usa Cloudflare Turnstile. Se resuelve con SeleniumBase en modo UC, que "
    "hace clic con el mouse real: el equipo no se puede usar mientras corre y no funciona en modo "
    "oculto. El tiempo del captcha cuenta dentro del tiempo de descarga.",
    "Los selectores dependen del HTML actual de la DIAN; si cambian, hay que ajustarlos.",
    "Extracción: los PDF de la DIAN traen capa de texto, que se lee con pdfplumber de forma exacta "
    "(~0,35 s por PDF). Si un PDF es solo imagen (escaneado), se usa OCR con Tesseract a 300 dpi "
    "(~3,6 s por PDF); requiere instalar Tesseract en Windows (TESSERACT_RUTA en .env).",
    "Exactitud medida del OCR frente a la capa de texto (10 facturas, 21 productos): número de "
    "factura, fecha, NIT y precio unitario 100 %, cantidad 95,2 %, código 66,7 % y descripción "
    "87,9 % de similitud (47,6 % exacta): el OCR confunde caracteres (S/5, 8/B, O/0) y las "
    "descripciones cortadas a mitad de palabra.",
    "Los PDF vienen con contraseña (NIT del emisor o del receptor consultado); se abren con el NIT "
    "del .env. La tabla se reconoce por los títulos de columna (Código, Descripción, Cantidad, "
    "Precio unitario): si la DIAN cambia el formato hay que ajustar los nombres.",
    "Se espera 0,5 s entre consultas (PAUSA_ENTRE_CONSULTAS) y se hacen hasta 5 intentos por CUFE (INTENTOS).",
    "Requiere Microsoft Edge o Google Chrome, conexión a internet y que el Excel de control no esté abierto.",
]
