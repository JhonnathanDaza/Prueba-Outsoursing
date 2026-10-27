"""PASO 1 - Leer el correo de la prueba (NIT del cuerpo + CUFEs del Word adjunto)."""

import email
import imaplib
import re
import zipfile
from email.header import decode_header
from html import unescape
from io import BytesIO

from config import CARPETA, CORREO_REMITENTE, EMAIL_CLAVE_APP, EMAIL_USUARIO, NIT_RESPALDO, log

def texto_docx(contenido):
    """Un .docx es un zip: el texto está en word/document.xml."""
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    xml = xml.replace("</w:p>", "\n")
    return unescape(re.sub(r"<[^>]+>", "", xml))


def buscar_cufes(texto):
    return list(dict.fromkeys(c.lower() for c in re.findall(r"\b[0-9a-fA-F]{96}\b", texto)))


def datos_de_mensaje(mensaje):
    """Saca el NIT (del cuerpo) y los CUFEs (del Word adjunto) de un correo."""
    texto = ""
    for parte in mensaje.walk():
        nombre = parte.get_filename()
        contenido = parte.get_payload(decode=True)
        if not contenido:
            continue
        if nombre:
            nombre = "".join(t.decode(c or "utf-8") if isinstance(t, bytes) else t
                             for t, c in decode_header(nombre))
            if nombre.lower().endswith(".docx"):
                texto += texto_docx(contenido) + "\n"
        elif parte.get_content_type() == "text/plain":
            texto += contenido.decode(parte.get_content_charset() or "utf-8", errors="replace")

    nit = re.search(r"NIT\s*:\s*([\d\.]{6,15})", texto, re.IGNORECASE)
    return (re.sub(r"\D", "", nit.group(1)) if nit else NIT_RESPALDO), buscar_cufes(texto)


def paso1_leer_correo():
    if not (EMAIL_USUARIO and EMAIL_CLAVE_APP):
        log("Paso 1 | Sin datos de correo en .env: uso cufes.txt y el NIT del .env")
        return NIT_RESPALDO, buscar_cufes((CARPETA / "cufes.txt").read_text(encoding="utf-8"))

    with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
        imap.login(EMAIL_USUARIO, EMAIL_CLAVE_APP)
        imap.select("INBOX", readonly=True)
        filtro = (["FROM", f'"{CORREO_REMITENTE}"'] if CORREO_REMITENTE else []) + ["SUBJECT", '"PRUEBA"']
        _, ids = imap.search(None, *filtro)
        if not ids[0]:
            raise LookupError(f"No encontré el correo de la prueba (remitente: {CORREO_REMITENTE or 'cualquiera'})")
        _, datos = imap.fetch(ids[0].split()[-1], "(RFC822)")
    nit, cufes = datos_de_mensaje(email.message_from_bytes(datos[0][1]))
    log(f"Paso 1 | Correo leído: NIT {nit}, {len(cufes)} CUFEs")
    return nit, cufes
