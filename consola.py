"""
CONSOLA DE OPERACIÓN - RPA DIAN

Herramienta de consola para operar el robot sin abrir el código:

    python consola.py relanzar                 Relanza el robot (solo los CUFEs marcados con X,
                                               o todos si no hay pendientes)
    python consola.py relanzar --todo          Relanza todos los CUFEs desde cero
    python consola.py relanzar --cufe 4 7      Marca con X los CUFEs 4 y 7 (N° o CUFE) y los relanza
    python consola.py relanzar --sin-correo    No envía el correo al final

    python consola.py informe                  Informe de la última ejecución del bot
    python consola.py informe --guardar        ... y lo guarda en salida/informe_ejecucion.txt
    python consola.py informe --json           ... en formato JSON (para otros sistemas)

    python consola.py estado                   Valida si la plataforma está operativa
                                               (DIAN, internet, Chrome, librerías, Excel, carpetas)

    python consola.py                          Menú interactivo

Códigos de salida: 0 = todo bien, 1 = hay fallas (útil para tareas programadas).
"""

import argparse
import csv
import importlib.util
import json
import re
import shutil
import socket
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

from config import (
    CARPETA, DESCARGAS, EMAIL_CLAVE_APP, EMAIL_USUARIO, EXCEL, SALIDA, URL_DIAN,
)

LOG = CARPETA / "rpa.log"
LINEA = "-" * 72


# =====================================================================
# Utilidades de presentación
# =====================================================================
def titulo(texto):
    print(f"\n{LINEA}\n {texto}\n{LINEA}")


def fila(estado, nombre, detalle=""):
    marca = {"OK": "[ OK ]", "AVISO": "[AVISO]", "FALLA": "[FALLA]"}[estado]
    print(f" {marca:8} {nombre:32} {detalle}")


def leer_excel():
    """(nit, filas) del Excel de control. filas: [N, CUFE, PROCESADO, NO PROCESADO, OBS, FECHA, REPROCESAR]."""
    from paso2_excel import leer_excel_existente
    return leer_excel_existente()


def marcado(valor):
    from paso2_excel import VALORES_MARCA
    return str(valor or "").strip().lower() in VALORES_MARCA


# =====================================================================
# 1. RELANZAR
# =====================================================================
def marcar_para_reproceso(identificadores):
    """Marca con X en REPROCESAR los CUFEs indicados (por N° o por CUFE completo/inicio)."""
    from paso2_excel import MARCA_REPROCESAR, ExcelControl
    nit, filas = leer_excel()
    if not filas:
        print(f" No existe {EXCEL.name} o no tiene CUFEs; ejecuta primero el robot completo.")
        return False
    marcados = []
    for ident in identificadores:
        ident = ident.strip().lower()
        encontrados = [f for f in filas if str(f[0]) == ident or f[1].lower().startswith(ident)]
        if not encontrados:
            print(f" [AVISO] No encontré el CUFE '{ident}' en el Excel; lo omito.")
            continue
        for f in encontrados:
            f[6] = MARCA_REPROCESAR
            marcados.append(f"{f[0]} ({f[1][:12]}...)")
    if not marcados:
        return False
    ExcelControl(nit, filas=filas)  # reescribe el Excel con las marcas
    print(f" Marcados con X para reprocesar: {', '.join(marcados)}")
    return True


def comando_relanzar(args):
    titulo("RELANZAMIENTO DEL ROBOT")
    if args.cufe:
        if not marcar_para_reproceso(args.cufe):
            return 1
    _, filas = leer_excel()
    pendientes = [f for f in filas if marcado(f[6])]
    if args.todo:
        print(" Modo: TODOS los CUFEs desde cero (--todo)")
    elif pendientes:
        print(f" Modo: REPROCESO de {len(pendientes)} CUFE(s) marcados con X:")
        for f in pendientes:
            print(f"   - N° {f[0]}  {f[1][:20]}...  {f[4] or ''}")
    else:
        print(" Modo: ejecución completa (no hay CUFEs marcados con X)")
    print(" Mientras corre no muevas el mouse: el robot lo usa para la casilla de Cloudflare.\n")

    comando = [sys.executable, str(CARPETA / "lanzador.py")]
    if args.todo:
        comando.append("--todo")
    if args.sin_correo:
        comando.append("--sin-correo")
    inicio = time.time()
    codigo = subprocess.call(comando, cwd=CARPETA)
    print(f"\n Relanzamiento terminado en {time.time() - inicio:.1f} s (código {codigo}).")
    print(" Revisa el resultado con:  python consola.py informe")
    return codigo


# =====================================================================
# 2. INFORME DE EJECUCIÓN
# =====================================================================
def ultima_ejecucion_log():
    """Líneas del log desde el último 'Lanzador | Inicio del robot'."""
    if not LOG.exists():
        return []
    lineas = LOG.read_text(encoding="utf-8", errors="replace").splitlines()
    inicio = max((i for i, l in enumerate(lineas) if "Lanzador | Inicio del robot" in l), default=None)
    if inicio is None:  # log anterior al lanzador
        inicio = max((i for i, l in enumerate(lineas) if "Paso 1 |" in l), default=0)
    return lineas[inicio:]


def leer_csv(ruta):
    if not ruta.exists():
        return []
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def datos_informe():
    nit, filas = leer_excel()
    exitosos = [f for f in filas if f[2] == "Exitoso"]
    fallidos = [f for f in filas if f[3] == "No exitoso"]
    pendientes = [f for f in filas if marcado(f[6])]
    sin_procesar = [f for f in filas if not f[2] and not f[3]]

    log = ultima_ejecucion_log()
    hora = lambda l: l.split(" | ")[0] if " | " in l else ""
    metricas = leer_csv(SALIDA / "metricas_descarga.csv")
    extraccion = leer_csv(SALIDA / "metricas_extraccion.csv")
    errores = [l for l in log if "Error en el intento" in l]
    modo = "reproceso" if any("Modo REPROCESO" in l for l in log) else "completa"
    tiempo_total = next((l.split(":", 1)[1].strip() for l in log if l.startswith("Tiempo total:")), "")
    fecha_log = datetime.fromtimestamp(LOG.stat().st_mtime).strftime("%Y-%m-%d") if LOG.exists() else ""
    pdfs = sorted(DESCARGAS.glob("*.pdf")) if DESCARGAS.exists() else []
    total = len(filas)
    return {
        "generado": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
        "ultima_ejecucion": {
            "fecha": fecha_log,
            "inicio": hora(log[0]) if log else "",
            "fin": hora(next((l for l in reversed(log) if " | " in l), "")),
            "terminada": any("Lanzador | Fin del robot" in l for l in log),
            "modo": modo,
            "cufes_consultados": len(metricas),
            "intentos_totales": sum(int(m.get("intentos") or 0) for m in metricas),
            "errores_en_intentos": len(errores),
            "tiempo_total": tiempo_total,
        },
        "excel_control": {
            "nit": nit,
            "total_cufes": total,
            "exitosos": len(exitosos),
            "no_exitosos": len(fallidos),
            "sin_procesar": len(sin_procesar),
            "por_reprocesar": len(pendientes),
            "porcentaje_exito": round(100 * len(exitosos) / total, 1) if total else 0.0,
        },
        "extraccion": {
            "pdfs_descargados": len(pdfs),
            "pdfs_con_datos_completos": sum(e.get("estado") == "OK" for e in extraccion),
            "csv_por_factura": len(list((SALIDA / "csv_por_factura").glob("*.csv")))
            if (SALIDA / "csv_por_factura").exists() else 0,
        },
        "pendientes_reproceso": [{"n": f[0], "cufe": f[1], "observacion": f[4]} for f in pendientes],
        "ultimos_errores": errores[-5:],
        "detalle": [{"n": f[0], "cufe": f[1], "estado": f[2] or f[3] or "Sin procesar",
                     "observacion": f[4], "fecha": f[5], "reprocesar": f[6]} for f in filas],
    }


def texto_informe(d):
    u, x, e = d["ultima_ejecucion"], d["excel_control"], d["extraccion"]
    out = [
        "INFORME DE EJECUCIÓN - RPA DIAN",
        f"Generado: {d['generado']}",
        "",
        "ÚLTIMA EJECUCIÓN",
        f"  Fecha / hora:          {u['fecha']}  {u['inicio']} -> {u['fin']}",
        f"  Estado:                {'Terminada' if u['terminada'] else 'Incompleta o en curso'}",
        f"  Modo:                  {u['modo']}",
        f"  CUFEs consultados:     {u['cufes_consultados']}",
        f"  Intentos totales:      {u['intentos_totales']}  (errores en intentos: {u['errores_en_intentos']})",
        f"  Tiempo total:          {u['tiempo_total'] or 'sin dato'}",
        "",
        "EXCEL DE CONTROL",
        f"  NIT:                   {x['nit']}",
        f"  Total CUFEs:           {x['total_cufes']}",
        f"  Exitosos:              {x['exitosos']}",
        f"  No exitosos:           {x['no_exitosos']}",
        f"  Sin procesar:          {x['sin_procesar']}",
        f"  Por reprocesar (X):    {x['por_reprocesar']}",
        f"  % de éxito:            {x['porcentaje_exito']}%",
        "",
        "EXTRACCIÓN",
        f"  PDFs descargados:      {e['pdfs_descargados']}",
        f"  PDFs con datos OK:     {e['pdfs_con_datos_completos']}",
        f"  CSV por factura:       {e['csv_por_factura']}",
        "",
        "DETALLE",
        f"  {'N°':>3}  {'ESTADO':13} {'REPROC.':7} {'FECHA':19}  CUFE",
    ]
    for f in d["detalle"]:
        out.append(f"  {f['n']:>3}  {f['estado']:13} {f['reprocesar'] or '':7} {f['fecha'] or '':19}  {f['cufe'][:24]}...")
    if d["pendientes_reproceso"]:
        out += ["", "PENDIENTES DE REPROCESO (relanzar con: python consola.py relanzar)"]
        out += [f"  N° {p['n']}: {p['observacion']}" for p in d["pendientes_reproceso"]]
    if d["ultimos_errores"]:
        out += ["", "ÚLTIMOS ERRORES DEL LOG"] + [f"  {l}" for l in d["ultimos_errores"]]
    return "\n".join(out)


def comando_informe(args):
    d = datos_informe()
    if not d["detalle"]:
        print(f" No hay datos: no existe {EXCEL.name}. Ejecuta primero el robot.")
        return 1
    if args.json:
        print(json.dumps(d, ensure_ascii=False, indent=2))
    else:
        titulo("INFORME DE EJECUCIÓN")
        print(texto_informe(d))
    if args.guardar:
        SALIDA.mkdir(exist_ok=True)
        (SALIDA / "informe_ejecucion.txt").write_text(texto_informe(d), encoding="utf-8")
        (SALIDA / "informe_ejecucion.json").write_text(json.dumps(d, ensure_ascii=False, indent=2),
                                                       encoding="utf-8")
        print(f"\n Guardado en {SALIDA / 'informe_ejecucion.txt'} (y .json)")
    return 1 if d["excel_control"]["por_reprocesar"] else 0


# =====================================================================
# 3. ESTADO DE LA PLATAFORMA
# =====================================================================
def revisar_dian():
    """La página de la DIAN responde? (Cloudflare puede devolver 403 con reto: igual está en línea)."""
    host = urllib.parse.urlparse(URL_DIAN).hostname
    try:
        t = time.time()
        socket.getaddrinfo(host, 443)
        dns_ms = (time.time() - t) * 1000
    except OSError as e:
        return "FALLA", f"no resuelve DNS de {host} ({e.__class__.__name__})"
    pedido = urllib.request.Request(URL_DIAN, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
        "Accept-Language": "es-CO,es;q=0.9"})
    t = time.time()
    try:
        with urllib.request.urlopen(pedido, timeout=20, context=ssl.create_default_context()) as r:
            codigo, cuerpo = r.status, r.read(200_000).decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        codigo, cuerpo = e.code, e.read(200_000).decode("utf-8", "replace") if e.fp else ""
    except Exception as e:
        return "FALLA", f"sin respuesta ({e.__class__.__name__}: {str(e)[:60]})"
    ms = (time.time() - t) * 1000
    detalle = f"HTTP {codigo} en {ms:.0f} ms (DNS {dns_ms:.0f} ms)"
    if codigo == 200 and ("SearchDocument" in cuerpo or "cf-turnstile" in cuerpo or "DIAN" in cuerpo):
        return ("AVISO" if ms > 8000 else "OK"), detalle + (" - lenta" if ms > 8000 else " - formulario disponible")
    if codigo in (403, 503) and ("cloudflare" in cuerpo.lower() or "challenge" in cuerpo.lower()):
        return "AVISO", detalle + " - en línea, Cloudflare pide verificación (normal para el robot)"
    if codigo >= 500:
        return "FALLA", detalle + " - el servidor de la DIAN está con errores"
    return "AVISO", detalle + " - respuesta inesperada"


def revisar_internet():
    try:
        socket.create_connection(("1.1.1.1", 443), timeout=5).close()
        return "OK", "conexión saliente disponible"
    except OSError:
        return "FALLA", "sin conexión a internet"


def revisar_chrome():
    from paso3_navegador import MODO_NAVEGADOR, NOMBRE_NAVEGADOR, buscar_navegador
    try:
        return "OK", f"{NOMBRE_NAVEGADOR} en {buscar_navegador()} (modo {MODO_NAVEGADOR})"
    except FileNotFoundError as e:
        return "FALLA", str(e)


def revisar_libreria(nombre, obligatoria=True):
    if importlib.util.find_spec(nombre):
        return "OK", "instalada"
    return ("FALLA" if obligatoria else "AVISO"), \
        "no instalada" + ("" if obligatoria else " (el robot la instala sola al usarla)")


def revisar_excel():
    if not EXCEL.exists():
        return "AVISO", "no existe todavía (se crea en la primera ejecución)"
    try:
        with open(EXCEL, "r+b"):
            pass
    except PermissionError:
        return "FALLA", "está abierto en Excel: ciérralo antes de ejecutar el robot"
    nit, filas = leer_excel()
    pend = sum(marcado(f[6]) for f in filas)
    return "OK", f"{len(filas)} CUFEs, {pend} por reprocesar"


def revisar_carpetas():
    problemas = []
    for carpeta in (DESCARGAS, SALIDA):
        try:
            carpeta.mkdir(exist_ok=True)
            prueba = carpeta / ".prueba_escritura"
            prueba.write_text("ok")
            prueba.unlink()
        except OSError:
            problemas.append(carpeta.name)
    libre = shutil.disk_usage(CARPETA).free / 1024 ** 3
    if problemas:
        return "FALLA", f"sin permiso de escritura en: {', '.join(problemas)}"
    return ("AVISO" if libre < 1 else "OK"), f"descargas/ y salida/ con escritura; {libre:.1f} GB libres"


def revisar_entradas():
    if EMAIL_USUARIO and EMAIL_CLAVE_APP:
        return "OK", f"lee el correo de {EMAIL_USUARIO}"
    cufes = CARPETA / "cufes.txt"
    if cufes.exists():
        n = len(re.findall(r"\b[0-9a-fA-F]{96}\b", cufes.read_text(encoding="utf-8")))
        return ("OK" if n else "FALLA"), f"sin correo configurado; usa cufes.txt ({n} CUFEs)"
    return "FALLA", "no hay correo en .env ni cufes.txt"


def revisar_correo_salida():
    if EMAIL_USUARIO and EMAIL_CLAVE_APP:
        return "OK", f"envía el Excel desde {EMAIL_USUARIO}"
    return "AVISO", "falta EMAIL_CLAVE_APP en .env: el paso 7 no envía el correo"


def revisar_ejecucion_en_curso():
    if LOG.exists() and time.time() - LOG.stat().st_mtime < 60:
        log = ultima_ejecucion_log()
        if log and not any("Lanzador | Fin del robot" in l for l in log):
            return "AVISO", "el robot parece estar corriendo ahora (el log cambió hace < 1 min)"
    return "OK", "no hay una ejecución en curso"


def comando_estado(args):
    titulo(f"ESTADO DE LA PLATAFORMA - {datetime.now():%Y-%m-%d %H:%M:%S}")
    revisiones = [
        ("Internet", revisar_internet),
        ("Portal DIAN (catálogo VPFE)", revisar_dian),
        ("Navegador", revisar_chrome),
        ("Librería selenium", lambda: revisar_libreria("selenium")),
        ("Librería seleniumbase", lambda: revisar_libreria("seleniumbase", obligatoria=False)),
        ("Librería pyautogui", lambda: revisar_libreria("pyautogui", obligatoria=False)),
        ("Excel de control", revisar_excel),
        ("Carpetas de trabajo", revisar_carpetas),
        ("Entrada de CUFEs", revisar_entradas),
        ("Correo de salida", revisar_correo_salida),
        ("Ejecución en curso", revisar_ejecucion_en_curso),
    ]
    resultados = []
    for nombre, revisar in revisiones:
        try:
            estado, detalle = revisar()
        except Exception as e:
            estado, detalle = "FALLA", f"{type(e).__name__}: {str(e)[:80]}"
        fila(estado, nombre, detalle)
        resultados.append({"revision": nombre, "estado": estado, "detalle": detalle})

    fallas = [r for r in resultados if r["estado"] == "FALLA"]
    avisos = [r for r in resultados if r["estado"] == "AVISO"]
    print(LINEA)
    if fallas:
        print(f" PLATAFORMA NO OPERATIVA: {len(fallas)} falla(s). Corrige lo marcado [FALLA].")
    elif avisos:
        print(f" PLATAFORMA OPERATIVA con {len(avisos)} aviso(s).")
    else:
        print(" PLATAFORMA OPERATIVA.")
    if args.json:
        print(json.dumps({"operativa": not fallas, "revisiones": resultados}, ensure_ascii=False, indent=2))
    return 1 if fallas else 0


# =====================================================================
# MENÚ INTERACTIVO Y ARGUMENTOS
# =====================================================================
def menu():
    opciones = {
        "1": ("Validar estado de la plataforma", lambda: comando_estado(argparse.Namespace(json=False))),
        "2": ("Informe de la última ejecución", lambda: comando_informe(
            argparse.Namespace(json=False, guardar=True))),
        "3": ("Relanzar (pendientes con X o completo)", lambda: comando_relanzar(
            argparse.Namespace(todo=False, cufe=None, sin_correo=False))),
        "4": ("Relanzar CUFEs específicos", lambda: comando_relanzar(argparse.Namespace(
            todo=False, sin_correo=False,
            cufe=input(" N° o inicio del CUFE (separados por espacio): ").split()))),
        "5": ("Relanzar todos desde cero", lambda: comando_relanzar(
            argparse.Namespace(todo=True, cufe=None, sin_correo=False))),
    }
    while True:
        titulo("CONSOLA RPA DIAN")
        for k, (texto, _) in opciones.items():
            print(f"  {k}. {texto}")
        print("  0. Salir")
        eleccion = input("\n Opción: ").strip()
        if eleccion == "0":
            return 0
        if eleccion in opciones:
            opciones[eleccion][1]()
            input("\n Enter para volver al menú...")


def main():
    parser = argparse.ArgumentParser(
        prog="consola.py", description="Consola de operación del RPA DIAN",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Sin comando abre el menú interactivo.")
    sub = parser.add_subparsers(dest="comando")

    p = sub.add_parser("relanzar", help="Relanza el robot (pendientes con X, CUFEs puntuales o todo)")
    p.add_argument("--todo", action="store_true", help="procesa todos los CUFEs desde cero")
    p.add_argument("--cufe", nargs="+", metavar="ID", help="N° o CUFE (o su inicio) a marcar con X y relanzar")
    p.add_argument("--sin-correo", action="store_true", help="no envía el correo al final")
    p.set_defaults(func=comando_relanzar)

    p = sub.add_parser("informe", help="Informe de la última ejecución del bot")
    p.add_argument("--guardar", action="store_true", help="guarda salida/informe_ejecucion.txt y .json")
    p.add_argument("--json", action="store_true", help="muestra el informe en JSON")
    p.set_defaults(func=comando_informe)

    p = sub.add_parser("estado", help="Valida si la plataforma está operativa")
    p.add_argument("--json", action="store_true", help="agrega el resultado en JSON")
    p.set_defaults(func=comando_estado)

    args = parser.parse_args()
    if not args.comando:
        return menu()
    return args.func(args)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n Cancelado.")
        sys.exit(130)
