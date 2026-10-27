"""PASO 2 - Crear el Excel de control (hojas CUFES y NIT). El .xlsx se arma a mano (zip con XML).

Reproceso: si un CUFE falla, su columna REPROCESAR queda marcada con "X". En la siguiente ejecución,
si el Excel tiene CUFEs marcados con X en REPROCESAR, el robot procesa SOLO esos (el resto se
conserva). También se puede marcar con X a mano cualquier fila. Con --todo se procesan todos de nuevo.
"""

import re
import time
import zipfile
from datetime import datetime
from xml.etree import ElementTree
from xml.sax.saxutils import escape

from config import CORREO_REMITENTE, EXCEL, log

ESTILO_NORMAL, ESTILO_TITULO, ESTILO_VERDE, ESTILO_ROJO, ESTILO_AMARILLO = 0, 1, 2, 3, 4
MARCA_REPROCESAR = "X"
# Se acepta X (también x, SI o 1 por si alguien lo escribe distinto a mano)
VALORES_MARCA = {"x", "si", "sí", "1", "true", "verdadero"}

XLSX_ESTILOS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

XLSX_FIJOS = {
    "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
    "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
    "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="CUFES" sheetId="1" r:id="rId1"/><sheet name="NIT" sheetId="2" r:id="rId2"/></sheets>
</workbook>""",
    "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
    "xl/styles.xml": XLSX_ESTILOS,
}


def xml_hoja(filas, anchos, estilos):
    """filas: lista de listas. estilos: {(fila, columna): estilo}. La fila 1 va en negrita."""
    columnas = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(anchos, 1))
    xml_filas = []
    for r, fila in enumerate(filas, 1):
        celdas = ""
        for c, valor in enumerate(fila, 1):
            estilo = estilos.get((r, c), ESTILO_TITULO if r == 1 else ESTILO_NORMAL)
            if valor in ("", None) and not estilo:
                continue
            s = f' s="{estilo}"' if estilo else ""
            ref = f"{chr(64 + c)}{r}"
            if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                celdas += f'<c r="{ref}"{s}><v>{valor}</v></c>'
            else:
                celdas += f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(valor or ""))}</t></is></c>'
        xml_filas.append(f'<row r="{r}">{celdas}</row>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
            'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{columnas}</cols><sheetData>{"".join(xml_filas)}</sheetData></worksheet>')


NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def leer_hoja(z, nombre, compartidos):
    """Lee una hoja del .xlsx como lista de filas (sirve para el Excel del robot y para uno
    guardado desde Excel, que usa textos compartidos)."""
    filas = {}
    for c in ElementTree.fromstring(z.read(nombre)).iterfind(".//m:c", NS):
        ref = re.match(r"([A-Z]+)(\d+)", c.get("r"))
        col = ord(ref.group(1)[-1]) - 64
        fila = int(ref.group(2))
        if c.get("t") == "inlineStr":
            valor = "".join(t.text or "" for t in c.iterfind(".//m:t", NS))
        else:
            v = c.find("m:v", NS)
            valor = v.text if v is not None else ""
            if c.get("t") == "s":
                valor = compartidos[int(valor)]
        filas.setdefault(fila, {})[col] = (valor or "").strip()
    return [[filas[f].get(c, "") for c in range(1, max(filas[f]) + 1)] for f in sorted(filas)]


def leer_excel_existente():
    """Devuelve (nit, filas) del Excel de control, o (None, []) si no existe o no se puede leer."""
    if not EXCEL.exists():
        return None, []
    try:
        with zipfile.ZipFile(EXCEL) as z:
            compartidos = []
            if "xl/sharedStrings.xml" in z.namelist():
                for si in ElementTree.fromstring(z.read("xl/sharedStrings.xml")).iterfind("m:si", NS):
                    compartidos.append("".join(t.text or "" for t in si.iterfind(".//m:t", NS)))
            cufes = leer_hoja(z, "xl/worksheets/sheet1.xml", compartidos)
            hoja_nit = leer_hoja(z, "xl/worksheets/sheet2.xml", compartidos)
    except Exception as e:
        log(f"Paso 2 | No pude leer {EXCEL.name} ({type(e).__name__}); se crea uno nuevo")
        return None, []
    nit = hoja_nit[1][0] if len(hoja_nit) > 1 and hoja_nit[1] else None
    filas = []
    for fila in cufes[1:]:  # la fila 1 son los títulos
        fila = (fila + [""] * 7)[:7]
        if re.fullmatch(r"[0-9a-fA-F]{96}", fila[1]):
            fila[0] = len(filas) + 1
            filas.append(fila)
    return nit, filas


class ExcelControl:
    """Guarda el estado en memoria y reescribe el .xlsx cada vez que algo cambia."""
    COLUMNAS = ["N°", "CUFE", "PROCESADO", "NO PROCESADO", "OBSERVACIÓN", "FECHA PROCESO", "REPROCESAR"]

    def __init__(self, nit, cufes=(), filas=None):
        """Excel nuevo con los CUFEs, o (filas=...) el contenido de un Excel existente."""
        self.nit = nit
        self.filas = filas if filas is not None else \
            [[n, cufe, "", "", "", "", ""] for n, cufe in enumerate(cufes, start=1)]
        self.guardar()

    def pendientes(self):
        """[(n, cufe)] de las filas marcadas para reprocesar."""
        return [(f[0], f[1]) for f in self.filas if str(f[6]).strip().lower() in VALORES_MARCA]

    def marcar(self, n, exitoso, observacion):
        """n = número del registro (1 = primer CUFE). Si falla, queda marcado con X en REPROCESAR."""
        fila = self.filas[n - 1]
        fila[2:7] = (["Exitoso", ""] if exitoso else ["", "No exitoso"]) + \
                    [observacion, f"{datetime.now():%Y-%m-%d %H:%M:%S}", "" if exitoso else MARCA_REPROCESAR]
        self.guardar()

    def estilos(self):
        """Colores según el estado de cada fila (verde, rojo y amarillo para REPROCESAR)."""
        estilos = {}
        for r, f in enumerate(self.filas, start=2):
            if f[2]:
                estilos[(r, 3)] = ESTILO_VERDE
            if f[3]:
                estilos[(r, 4)] = ESTILO_ROJO
            if str(f[6]).strip().lower() in VALORES_MARCA:
                estilos[(r, 7)] = ESTILO_AMARILLO
        return estilos

    def guardar(self):
        hoja_cufes = xml_hoja([self.COLUMNAS] + self.filas, [6, 100, 14, 16, 45, 20, 13], self.estilos())
        hoja_nit = xml_hoja([["NIT", "ORIGEN", "FECHA"],
                             [self.nit, f"Correo de {CORREO_REMITENTE}" if CORREO_REMITENTE else "Correo de la prueba / cufes.txt", f"{datetime.now():%Y-%m-%d %H:%M}"]],
                            [14, 45, 18], {})
        for intento in range(5):
            try:
                with zipfile.ZipFile(EXCEL, "w", zipfile.ZIP_DEFLATED) as z:
                    for nombre, contenido in XLSX_FIJOS.items():
                        z.writestr(nombre, contenido)
                    z.writestr("xl/worksheets/sheet1.xml", hoja_cufes)
                    z.writestr("xl/worksheets/sheet2.xml", hoja_nit)
                return
            except PermissionError:
                log("Excel | El archivo está abierto en Excel; ciérralo para continuar")
                time.sleep(3)
        raise PermissionError(f"No se pudo guardar {EXCEL.name}")


def paso2_crear_excel(nit, cufes, todo=False):
    """Devuelve (excel, pendientes). pendientes = [(n, cufe)] que el paso 4 debe procesar.

    - Si el Excel de control ya existe y tiene CUFEs marcados con X en REPROCESAR (y no se pidió --todo),
      se conserva y solo se procesan esos.
    - Si no, se crea un Excel nuevo y se procesan todos los CUFEs.
    """
    if not todo:
        nit_excel, filas = leer_excel_existente()
        if filas:
            excel = ExcelControl(nit_excel or nit, filas=filas)
            pendientes = excel.pendientes()
            if pendientes:
                log(f"Paso 2 | Modo REPROCESO: {len(pendientes)} de {len(filas)} CUFEs marcados "
                    f"para reprocesar en {EXCEL.name}")
                return excel, pendientes
    excel = ExcelControl(nit, cufes)
    log(f"Paso 2 | Excel creado: {len(cufes)} registros, NIT {nit}")
    return excel, list(enumerate(cufes, start=1))
