"""Configuración compartida del robot: .env, rutas, parámetros y utilidades (log, CSV)."""

import csv
import os
import unicodedata
from datetime import datetime
from pathlib import Path

CARPETA = Path(__file__).resolve().parent


def cargar_env():
    """Lee el archivo .env (líneas CLAVE=valor)."""
    ruta = CARPETA / ".env"
    if ruta.exists():
        for linea in ruta.read_text(encoding="utf-8").splitlines():
            if "=" in linea and not linea.strip().startswith("#"):
                clave, valor = linea.split("=", 1)
                os.environ.setdefault(clave.strip(), valor.strip())


cargar_env()
EMAIL_USUARIO = os.getenv("EMAIL_USUARIO", "")
EMAIL_CLAVE_APP = os.getenv("EMAIL_CLAVE_APP", "").replace(" ", "")
CORREO_DESTINO = os.getenv("CORREO_DESTINO") or EMAIL_USUARIO
CORREO_REMITENTE = os.getenv("CORREO_REMITENTE", "")  # remitente del correo de la prueba (vacío = cualquiera)
NIT_RESPALDO = os.getenv("NIT", "900156264")
API_2CAPTCHA = os.getenv("TWOCAPTCHA_API_KEY", "")  # opcional

URL_DIAN = "https://catalogo-vpfe.dian.gov.co/User/SearchDocument"
DESCARGAS = CARPETA / "descargas"
SALIDA = CARPETA / "salida"
EXCEL = CARPETA / "control_facturas.xlsx"
INTENTOS = int(os.getenv("INTENTOS", "5"))  # intentos por CUFE antes de marcarlo No exitoso
ESPERA_CAPTCHA = int(os.getenv("ESPERA_CAPTCHA", "90"))  # segundos para resolver el captcha a mano
RECARGAS_CAPTCHA = int(os.getenv("RECARGAS_CAPTCHA", "5"))  # recargas de la página si no se resuelve
PAUSA_ENTRE_CONSULTAS = float(os.getenv("PAUSA_ENTRE_CONSULTAS", "0.5"))  # segundos entre facturas
RECONEXION_UC = float(os.getenv("RECONEXION_UC", "2.5"))  # segundos con el driver desconectado al abrir
PAUSA_VERIFICACION = float(os.getenv("PAUSA_VERIFICACION", "3"))  # s sin driver tras clic en la casilla

def log(texto):
    linea = f"{datetime.now():%H:%M:%S} | {texto}"
    print(linea)
    with open(CARPETA / "rpa.log", "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def sin_tildes(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()

def guardar_csv(ruta, filas, campos):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        escritor.writeheader()
        escritor.writerows(filas)
