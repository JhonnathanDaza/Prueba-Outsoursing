"""PASO 4 - Consultar cada CUFE en la DIAN, descargar el PDF y marcar el Excel."""

import json
import os
import shutil
import time
import urllib.parse
import urllib.request
from pathlib import Path

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from config import (
    API_2CAPTCHA, CARPETA, DESCARGAS, ESPERA_CAPTCHA, INTENTOS, PAUSA_ENTRE_CONSULTAS,
    RECARGAS_CAPTCHA, RECONEXION_UC, SALIDA, URL_DIAN, log, sin_tildes,
)
from paso3_navegador import clic_captcha_uc, es_uc, funcion_uc, pausa_desconectado

# si = primero marca la casilla con JavaScript (checked = true); no = va directo al clic con el mouse
CASILLA_JS = os.getenv("CASILLA_JS", "si").lower() in ("si", "sí", "1", "true")

JS_DETECTAR_CAPTCHA = """
const el = document.querySelector('[data-sitekey]');
if (el) return [el.className.includes('turnstile') ? 'turnstile' : 'recaptcha', el.dataset.sitekey];
for (const s of document.scripts) {
  const m = (s.src || '').match(/recaptcha\\/(?:api|enterprise)\\.js\\?render=([^&]+)/);
  if (m && m[1] !== 'explicit') return ['recaptcha_v3', m[1]];
}
for (const f of document.querySelectorAll('iframe')) {
  const src = f.src || '';
  if (src.includes('challenges.cloudflare.com')) { const m = src.match(/0x4[\\w-]+/); return ['turnstile', m ? m[0] : null]; }
  if (src.includes('recaptcha')) { const m = src.match(/[?&]k=([^&]+)/); return ['recaptcha', m ? m[1] : null]; }
}
return [null, null];
"""
JS_TOKEN = """
// Puede haber varios campos (formulario de búsqueda y página de la factura): basta con uno lleno
for (const e of document.querySelectorAll(
    '[name="g-recaptcha-response"], [name="cf-turnstile-response"], [name$="turnstile-response"], [name="RecaptchaToken"]')) {
  if (e.value) return e.value;
}
try { if (window.turnstile) { const t = window.turnstile.getResponse(); if (t) return t; } } catch (e) {}
return '';
"""
JS_PONER_TOKEN = """
document.querySelectorAll('[name="g-recaptcha-response"], [name="cf-turnstile-response"], [name="RecaptchaToken"]')
  .forEach(e => { e.value = arguments[0]; e.innerHTML = arguments[0]; });
"""


def token_2captcha(tipo, sitekey, url):
    """Pide el token a 2Captcha (opcional, solo si hay API key en el .env)."""
    parametros = {"key": API_2CAPTCHA, "pageurl": url, "json": 1}
    if tipo == "turnstile":
        parametros.update(method="turnstile", sitekey=sitekey)
    else:
        parametros.update(method="userrecaptcha", googlekey=sitekey)
        if tipo == "recaptcha_v3":
            parametros.update(version="v3", min_score=0.7)
    with urllib.request.urlopen("https://2captcha.com/in.php?" + urllib.parse.urlencode(parametros)) as r:
        id_tarea = json.load(r)["request"]
    for _ in range(40):
        time.sleep(5)
        consulta = urllib.parse.urlencode({"key": API_2CAPTCHA, "action": "get", "id": id_tarea, "json": 1})
        with urllib.request.urlopen("https://2captcha.com/res.php?" + consulta) as r:
            respuesta = json.load(r)
        if respuesta["status"] == 1:
            return respuesta["request"]
    raise TimeoutError("2Captcha no respondió")


CASILLA_HUMANO = ("//input[@type='checkbox' and "
                  "(@aria-label='Verifique que es un ser humano' or contains(@aria-label, 'humano') "
                  "or contains(@aria-label, 'human'))]")


JS_MARCAR_CASILLA = """
// Cambia el estado de la casilla a marcado (checked = true) sin hacer clic
const casilla = arguments[0];
casilla.checked = true;
casilla.setAttribute('checked', 'checked');
casilla.setAttribute('aria-checked', 'true');
casilla.dispatchEvent(new Event('input', {bubbles: true}));
casilla.dispatchEvent(new Event('change', {bubbles: true}));
return casilla.checked;
"""


def marcar_casilla_humano(driver, espera_token=5):
    """Marca <input type="checkbox" aria-label="Verifique que es un ser humano">.

    La casilla vive dentro del iframe de Cloudflare (challenges.cloudflare.com).
    1) Busca la casilla en la página y en cada iframe de Cloudflare y cambia su
       estado a marcado con JavaScript (checked = true), sin hacer clic.
    2) Si así no aparece el token de verificación (o la casilla está en un shadow
       DOM cerrado que Selenium no ve), hace clic con el mouse sobre su posición.
    Devuelve True si pudo marcarla.
    """
    driver.switch_to.default_content()

    # 1) Cambiar el estado a marcado (se apaga con CASILLA_JS=no en el .env)
    marcada = False
    casillas = driver.find_elements(By.XPATH, CASILLA_HUMANO) if CASILLA_JS else []
    if not CASILLA_JS:
        pass
    elif casillas:
        marcada = driver.execute_script(JS_MARCAR_CASILLA, casillas[0])
    else:
        for iframe in driver.find_elements(By.CSS_SELECTOR, "iframe[src*='challenges.cloudflare.com']"):
            try:
                driver.switch_to.frame(iframe)
                casilla = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.XPATH, CASILLA_HUMANO)))
                marcada = driver.execute_script(JS_MARCAR_CASILLA, casilla)
            except Exception:
                pass
            finally:
                driver.switch_to.default_content()
            if marcada:
                break
    if marcada:
        log("Paso 4 | Casilla 'Verifique que es un ser humano' cambiada a marcado (checked = true)")
        try:
            WebDriverWait(driver, espera_token).until(lambda d: d.execute_script(JS_TOKEN))
            return True
        except TimeoutException:
            log("Paso 4 | Cloudflare no aceptó el cambio de estado: pruebo con clic")

    # 2) Respaldo: clic por posición; la casilla está a ~30 px del borde izquierdo del widget
    try:
        widget = WebDriverWait(driver, 10).until(
            lambda d: next((w for w in d.find_elements(
                By.CSS_SELECTOR, "[data-sitekey], .cf-turnstile, iframe[src*='challenges.cloudflare.com']")
                if w.size["width"] > 50 and w.size["height"] > 30), False))
        time.sleep(1.5)  # deja que el widget termine de dibujarse
        ActionChains(driver).move_to_element_with_offset(
            widget, -widget.size["width"] // 2 + 30, 0).pause(0.3).click().perform()
        log("Paso 4 | Clic en la casilla 'Verifique que es un ser humano' (por posición)")
        return True
    except Exception as e:
        log(f"Paso 4 | No pude hacer clic en la casilla: {type(e).__name__}")
        return False


def resolver_captcha(driver, espera_max=None):
    """Devuelve los segundos que tomó. Sin API key: espera a que lo resuelvas en el navegador.
    espera_max: tope de segundos para esperar el token (None = ESPERA_CAPTCHA / 25 s en UC)."""
    inicio = time.time()
    tipo, sitekey = driver.execute_script(JS_DETECTAR_CAPTCHA)
    if tipo is None:
        return 0
    log(f"Paso 4 | Captcha detectado: {tipo}")
    if API_2CAPTCHA and sitekey:
        driver.execute_script(JS_PONER_TOKEN, token_2captcha(tipo, sitekey, driver.current_url))
    elif tipo != "recaptcha_v3":  # el v3 es invisible: el navegador lo resuelve solo
        if driver.execute_script(JS_TOKEN):
            return round(time.time() - inicio, 1)
        # El widget de la DIAN está fijo en la esquina inferior derecha y es fácil no verlo
        driver.execute_script("const w = document.querySelector('[data-sitekey]');"
                              "if (w) w.scrollIntoView({block: 'center'});")
        if es_uc(driver):
            # Modo UC: SeleniumBase ubica el widget de Cloudflare y hace clic en la casilla
            # con el mouse real (PyAutoGUI). No se toca el iframe, para no delatar al robot.
            log("Paso 4 | Modo UC: clic en la casilla 'Verifique que es un ser humano' "
                "(no muevas el mouse unos segundos)")
            clic_captcha_uc(driver)
        else:
            # El RPA marca/hace clic en la casilla; si Cloudflare pide más, se puede marcar a mano
            marcar_casilla_humano(driver)
        # En modo UC, si en 25 s no hay token (p. ej. error 600010) se recarga y se reintenta
        espera = min(ESPERA_CAPTCHA, 25) if es_uc(driver) else ESPERA_CAPTCHA
        if espera_max:
            espera = min(espera, espera_max)
        log(f"Paso 4 | Esperando la verificación (máx. {espera} s). Si no avanza, marca la "
            f"casilla 'Verifique que es un ser humano' en la esquina inferior derecha del navegador")
        WebDriverWait(driver, espera, poll_frequency=0.25).until(
            lambda d: d.execute_script(JS_TOKEN),
            "No se resolvió el captcha a tiempo")
    return round(time.time() - inicio, 1)


# Carpetas donde puede caer el PDF (SeleniumBase usa downloaded_files/ por defecto)
CARPETAS_DESCARGA = list(dict.fromkeys(
    [DESCARGAS, CARPETA / "downloaded_files", Path.cwd() / "downloaded_files"]))


def pdfs_descargados():
    return {p for c in CARPETAS_DESCARGA if c.exists() for p in c.glob("*.pdf")}


def esperar_pdf(antes, segundos=60):
    fin = time.time() + segundos
    while time.time() < fin:
        nuevos = pdfs_descargados() - antes
        en_curso = [p for c in CARPETAS_DESCARGA if c.exists() for p in c.glob("*.crdownload")]
        if nuevos and not en_curso:
            return nuevos.pop()
        time.sleep(0.25)
    raise TimeoutError("El PDF no se descargó")


def abrir_pagina_dian(driver):
    if es_uc(driver):
        # Abre la página con chromedriver desconectado unos segundos para que Cloudflare
        # no detecte la automatización mientras carga
        funcion_uc(driver, "uc_open_with_reconnect")(URL_DIAN, RECONEXION_UC)
    else:
        driver.get(URL_DIAN)


def llenar_formulario(driver, cufe, nit):
    abrir_pagina_dian(driver)
    campo_cufe = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
    campo_cufe.clear()
    campo_cufe.send_keys(cufe)
    campo_nit = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'NIT')]")
    campo_nit.clear()
    campo_nit.send_keys(nit)


XPATH_DESCARGAR_PDF = "//a[contains(translate(., 'descargarpdf', 'DESCARGARPDF'), 'DESCARGAR PDF')]"
XPATH_ACEPTAR_AVISO = ("//button[translate(normalize-space(.), 'ACEPTR', 'aceptr')='aceptar'] | "
                       "//a[translate(normalize-space(.), 'ACEPTR', 'aceptr')='aceptar'] | "
                       "//input[@type='button' or @type='submit'][translate(@value, 'ACEPTR', 'aceptr')='aceptar']")


JS_TEXTO_AVISO = """
// Texto del aviso visible: se busca el botón "Aceptar" visible y se sube hasta el recuadro
// que contiene el mensaje (así no importa qué librería use la DIAN para el aviso)
const botones = [...document.querySelectorAll('button, a, input[type=button], input[type=submit]')]
  .filter(b => /^\\s*aceptar\\s*$/i.test(b.innerText || b.value || '') && b.offsetParent !== null);
for (const b of botones) {
  let el = b;
  for (let i = 0; i < 8 && el.parentElement; i++) {
    el = el.parentElement;
    const t = (el.innerText || '').trim();
    if (t.length > 40) return t;
  }
}
return '';
"""


def cerrar_aviso_contrasena(driver, segundos=3):
    """Cierra el aviso 'Este archivo contiene contraseña y corresponde al NIT...' (botón Aceptar)."""
    fin = time.time() + segundos
    while time.time() < fin:
        for boton in driver.find_elements(By.XPATH, XPATH_ACEPTAR_AVISO):
            if boton.is_displayed():
                driver.execute_script("arguments[0].click();", boton)
                time.sleep(0.3)
                return True
        time.sleep(0.25)
    return False


def descargar_factura(driver, cufe, nit):
    espera = WebDriverWait(driver, 30)

    # Si la verificación falla o no se completa, se recarga la página y se vuelve a pedir
    inicio = time.time()
    for vuelta in range(1, RECARGAS_CAPTCHA + 1):
        llenar_formulario(driver, cufe, nit)
        try:
            resolver_captcha(driver)
            break
        except TimeoutException:
            log(f"Paso 4 | Verificación no completada ({vuelta}/{RECARGAS_CAPTCHA}): recargo la página")
    else:
        raise TimeoutError(f"Captcha no resuelto tras {RECARGAS_CAPTCHA} recargas")
    # Clic por JavaScript: no se queda esperando a que la página termine de cargar
    boton_buscar = driver.find_element(By.XPATH, "//button[contains(., 'Buscar')]")
    driver.execute_script("arguments[0].click();", boton_buscar)
    log("Paso 4 | Captcha superado; consulta enviada")
    # La página de la factura trae otra casilla de Cloudflare: se deja cargar sin chromedriver
    pausa_desconectado(driver, RECONEXION_UC + 2)

    # Página de la factura: el enlace "Descargar PDF"
    espera.until(EC.presence_of_element_located((By.XPATH, XPATH_DESCARGAR_PDF)))
    # Lo primero al entrar a la página de la factura: descargar el PDF (descargar_pdf).
    # La casilla de Cloudflare solo se marca si la DIAN rechaza la descarga.
    t_captcha = round(time.time() - inicio, 1)

    destino = DESCARGAS / f"{cufe}.pdf"
    destino.unlink(missing_ok=True)
    shutil.move(str(descargar_pdf(driver)), str(destino))
    return t_captcha


JS_TOKEN_TURNSTILE = """
// Igual que getTurnstileValue() de la DIAN: el token de Cloudflare está en cf-turnstile-response
for (const e of document.querySelectorAll('input[name="cf-turnstile-response"]')) {
  if (e.value) return e.value.length;
}
return 0;
"""


def guardar_html_depuracion(driver, nombre):
    """Guarda el HTML de la página una vez por ejecución para revisar cómo está armada."""
    ruta = SALIDA / "depuracion" / f"{nombre}.html"
    if not ruta.exists():
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text(driver.page_source, encoding="utf-8")


def esperar_token(driver, segundos):
    """Espera a que Cloudflare deje el token en cf-turnstile-response. True si llegó."""
    fin = time.time() + segundos
    while time.time() < fin:
        if driver.execute_script(JS_TOKEN_TURNSTILE):
            return True
        time.sleep(0.25)
    return False


def descargar_pdf(driver, segundos=150):
    """Descarga el PDF desde la página de la factura.

    Cómo funciona la página de la DIAN (visto en su HTML):
      clic en 'Descargar PDF' -> aviso 'contiene contraseña' -> Aceptar -> si cf-turnstile-response
      tiene token, envía el formulario y descarga; si no, muestra 'campo de seguridad no completo'.
      Después de cada intento llama turnstile.reset(), así que hacer clic sin token lo empeora.

    Por eso el RPA:
      1. Espera el token de Cloudflare (si en 6 s no llega, marca la casilla con el mouse).
      2. Con token, hace UN clic en 'Descargar PDF' y acepta el aviso de contraseña.
      3. Espera el archivo. Si la DIAN rechaza, espera un token nuevo y repite.
    """
    antes = pdfs_descargados()
    fin = time.time() + segundos
    intento = 0
    while time.time() < fin:
        # 1. Token de Cloudflare
        if not esperar_token(driver, 6):
            log("Paso 4 | Cloudflare aún no valida la página de la factura: marco la casilla")
            if es_uc(driver):
                clic_captcha_uc(driver)
            if not esperar_token(driver, 15):
                # 'La verificación falló': se reinicia la casilla y se deja verificar sin chromedriver
                log("Paso 4 | La verificación de Cloudflare no se completó: reinicio la casilla")
                guardar_html_depuracion(driver, "pagina_factura_sin_token")
                driver.execute_script("try { turnstile.reset(); } catch (e) {}")
                pausa_desconectado(driver, RECONEXION_UC + 2)
                continue
        intento += 1
        log(f"Paso 4 | Casilla de Cloudflare lista; clic en 'Descargar PDF' ({intento})")

        # 2. Clic en 'Descargar PDF' y Aceptar en el aviso de contraseña
        enlace = driver.find_element(By.XPATH, XPATH_DESCARGAR_PDF)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();",
                              enlace)
        if cerrar_aviso_contrasena(driver, segundos=5):
            log("Paso 4 | Aviso 'el PDF tiene contraseña (NIT)' aceptado")

        # 3. Esperar el archivo (o un rechazo de la DIAN)
        limite = time.time() + 30
        while time.time() < limite:
            nuevos = pdfs_descargados() - antes
            en_curso = any(c.exists() and any(c.glob("*.crdownload")) for c in CARPETAS_DESCARGA)
            if nuevos and not en_curso:
                return nuevos.pop()
            texto = sin_tildes(driver.execute_script(JS_TEXTO_AVISO) or "")
            if "contrasena" in texto:
                cerrar_aviso_contrasena(driver, segundos=1)  # el aviso de contraseña se repitió
            elif "seguridad" in texto or "pdf" in texto:
                # 'campo de seguridad no completo' o 'No se pudo generar el PDF'
                cerrar_aviso_contrasena(driver, segundos=1)
                log(f"Paso 4 | La DIAN rechazó la descarga: {texto[:70]!r}; reintento con token nuevo")
                time.sleep(1)  # la DIAN reinicia la casilla (turnstile.reset)
                break
            time.sleep(0.25)
    raise TimeoutError("El PDF no se descargó")


def paso4_descargar_todo(driver, excel, nit, pendientes):
    """pendientes = [(n, cufe)]: n es la fila del CUFE en el Excel de control."""
    metricas = []
    total = len(pendientes)
    for i, (n, cufe) in enumerate(pendientes, start=1):
        inicio = time.time()
        fila = {"n": n, "cufe": cufe, "estado": "No exitoso", "intentos": 0,
                "t_captcha_s": 0, "t_total_s": 0, "observacion": ""}
        for intento in range(1, INTENTOS + 1):
            fila["intentos"] = intento
            try:
                log(f"Paso 4 | [{i}/{total}] CUFE {cufe[:12]}... (fila {n}) intento {intento}")
                fila["t_captcha_s"] = descargar_factura(driver, cufe, nit)
                fila["estado"], fila["observacion"] = "Exitoso", "PDF descargado"
                break
            except Exception as e:
                fila["observacion"] = f"{type(e).__name__}: {e}".split("\n")[0][:150]
                log(f"Paso 4 | Error en el intento {intento}: {fila['observacion']}")
                driver.save_screenshot(str(DESCARGAS / f"error_{n}_{intento}.png"))
        fila["t_total_s"] = round(time.time() - inicio, 1)
        excel.marcar(n, fila["estado"] == "Exitoso", fila["observacion"])
        log(f"Paso 4 | [{i}/{total}] {fila['estado']} - marcado en el Excel"
            + ("" if fila["estado"] == "Exitoso" else " (marcado con X en REPROCESAR)"))
        metricas.append(fila)
        time.sleep(PAUSA_ENTRE_CONSULTAS)
    return metricas
