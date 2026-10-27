"""PASO 7 - Enviar el Excel por correo."""

import smtplib
from datetime import datetime
from email.message import EmailMessage

from config import CORREO_DESTINO, EMAIL_CLAVE_APP, EMAIL_USUARIO, EXCEL, log

def paso7_enviar_correo(resumen):
    if not (EMAIL_USUARIO and EMAIL_CLAVE_APP):
        log("Paso 7 | Correo no enviado: faltan EMAIL_USUARIO o EMAIL_CLAVE_APP en .env")
        return
    msg = EmailMessage()
    msg["Subject"] = f"RPA DIAN - Resultado {datetime.now():%Y-%m-%d %H:%M}"
    msg["From"], msg["To"] = EMAIL_USUARIO, CORREO_DESTINO
    msg.set_content("Resultado del RPA:\n\n" + resumen + "\nSe adjunta el Excel de control.")
    msg.add_attachment(EXCEL.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=EXCEL.name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_USUARIO, EMAIL_CLAVE_APP)
        smtp.send_message(msg)
    log(f"Paso 7 | Excel enviado a {CORREO_DESTINO}")
