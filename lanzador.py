"""
LANZADOR - RPA DIAN. Ejecuta el robot en orden, un archivo por paso:

  1. paso1_correo.py      Leer el correo de la prueba (NIT del cuerpo + CUFEs del Word adjunto)
  2. paso2_excel.py       Crear el Excel de control (hojas CUFES y NIT)
  3. paso3_navegador.py   Abrir Chrome con SeleniumBase en modo UC (o Edge/Chrome con Selenium)
  4. paso4_descarga.py    Consultar cada CUFE en la DIAN, descargar el PDF y marcar el Excel
  5. paso5_extraccion.py  Extraer el texto de cada factura (capa de texto u OCR), un DataFrame por archivo -> CSV
  6. paso6_reporte.py     Mostrar y guardar tiempos y métricas
  7. paso7_envio.py       Enviar el Excel por correo

La configuración compartida (.env, rutas, log) está en config.py.

Ejecutar:
    python lanzador.py              # todo el proceso (o solo los CUFEs marcados con X en REPROCESAR)
    python lanzador.py --todo       # procesa todos los CUFEs aunque haya pendientes de reproceso
    python lanzador.py --sin-correo # no envía el correo al final

Reproceso: si un CUFE falla, queda marcado con X en la columna REPROCESAR de control_facturas.xlsx. La siguiente
ejecución procesa solo esos CUFEs. También se puede marcar con X a mano en esa columna.
"""

import sys
import time

from config import log
from paso1_correo import paso1_leer_correo
from paso2_excel import paso2_crear_excel
from paso3_navegador import cerrar_chrome, paso3_abrir_navegador
from paso4_descarga import paso4_descargar_todo
from paso5_extraccion import paso5_extraer_todo
from paso6_reporte import paso6_reporte
from paso7_envio import paso7_enviar_correo


def main():
    tiempos, t0 = {}, time.time()
    log("Lanzador | Inicio del robot")

    # Pasos 1 y 2: datos de la prueba y Excel de control
    t = time.time()
    nit, cufes = paso1_leer_correo()
    excel, pendientes = paso2_crear_excel(nit, cufes, todo="--todo" in sys.argv)
    tiempos["correo y Excel"] = time.time() - t

    # Paso 3: navegador (se cierra siempre al final, aunque algo falle)
    driver = paso3_abrir_navegador()
    try:
        # Paso 4: consulta en la DIAN y descarga de los PDF
        t = time.time()
        met_descarga = paso4_descargar_todo(driver, excel, nit, pendientes)
        tiempos["descarga"] = time.time() - t
    finally:
        cerrar_chrome(driver)

    # Paso 5: extracción del texto de las facturas -> un DataFrame por archivo -> CSV
    # (no necesita el navegador: lee los PDF con pdfplumber y, si son imágenes, con OCR)
    t = time.time()
    met_extraccion = paso5_extraer_todo(nit)
    tiempos["extracción"] = time.time() - t
    tiempos["total"] = time.time() - t0

    # Paso 6: tiempos y métricas
    resumen = paso6_reporte(tiempos, met_descarga, met_extraccion)
    por_reprocesar = excel.pendientes()
    if por_reprocesar:
        log(f"Lanzador | {len(por_reprocesar)} CUFE(s) quedaron marcados con X en REPROCESAR; "
            "vuelve a ejecutar el lanzador para procesar solo esos")

    # Paso 7: envío del Excel por correo (se omite con --sin-correo)
    if "--sin-correo" not in sys.argv:
        paso7_enviar_correo(resumen)
    log("Lanzador | Fin del robot")


if __name__ == "__main__":
    main()
