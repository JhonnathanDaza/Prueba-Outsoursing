"""PASO 3 - Abrir Chrome con SeleniumBase en modo UC (o Edge/Chrome con Selenium).
Incluye las utilidades del modo UC (desconectar el driver, clic real en la casilla de Cloudflare)."""

import json
import os
import shutil
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

from selenium import webdriver

from config import CARPETA, DESCARGAS, PAUSA_VERIFICACION, log

# Cómo se abre el navegador:
#   uc       -> Google Chrome con SeleniumBase en modo UC (undetected). Oculta las señales de
#               automatización y hace clic en la casilla de Cloudflare con el mouse real. (Recomendado)
#   normal   -> abre Edge/Chrome normal (sin el driver de Selenium) y Selenium se conecta a él.
#   selenium -> navegador lanzado por Selenium (más simple, pero Cloudflare suele rechazarlo)
MODO_NAVEGADOR = (os.getenv("MODO_NAVEGADOR") or os.getenv("MODO_CHROME") or "uc").lower()
# Navegador para los modos normal y selenium: edge o chrome. El modo uc siempre usa Chrome.
NAVEGADOR = "chrome" if MODO_NAVEGADOR == "uc" else os.getenv("NAVEGADOR", "edge").lower()
NAVEGADOR_RUTA = os.getenv("NAVEGADOR_RUTA") or os.getenv("CHROME_RUTA", "")  # vacío = se busca solo
PUERTO_DEPURACION = int(os.getenv("PUERTO_DEPURACION", "9222"))
PERFIL_NAVEGADOR = CARPETA / f"perfil_{NAVEGADOR}"  # perfil propio del RPA (no toca tu navegador personal)

RUTAS_NAVEGADOR = {
    "edge": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "microsoft-edge", "microsoft-edge-stable", "msedge",
    ],
    "chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "google-chrome", "google-chrome-stable", "chrome",
    ],
}
NOMBRE_NAVEGADOR = {"edge": "Microsoft Edge", "chrome": "Google Chrome"}.get(NAVEGADOR, NAVEGADOR)


def buscar_navegador():
    for ruta in [NAVEGADOR_RUTA] + RUTAS_NAVEGADOR.get(NAVEGADOR, []):
        ruta = ruta and (shutil.which(ruta) or ruta)
        if ruta and Path(ruta).exists():
            return ruta
    raise FileNotFoundError(f"No encontré {NOMBRE_NAVEGADOR}; pon su ruta en NAVEGADOR_RUTA en el .env")


def opciones_selenium():
    """Opciones y clase del driver de Selenium según el navegador elegido."""
    if NAVEGADOR == "edge":
        return webdriver.EdgeOptions(), webdriver.Edge
    return webdriver.ChromeOptions(), webdriver.Chrome


def preparar_perfil():
    """Deja configuradas las descargas en el perfil del RPA antes de abrir el navegador."""
    archivo = PERFIL_NAVEGADOR / "Default" / "Preferences"
    archivo.parent.mkdir(parents=True, exist_ok=True)
    try:
        prefs = json.loads(archivo.read_text(encoding="utf-8"))
    except Exception:
        prefs = {}
    prefs.setdefault("download", {}).update(
        default_directory=str(DESCARGAS), prompt_for_download=False, directory_upgrade=True)
    prefs.setdefault("savefile", {})["default_directory"] = str(DESCARGAS)
    prefs.setdefault("plugins", {})["always_open_pdf_externally"] = True  # descarga el PDF
    archivo.write_text(json.dumps(prefs), encoding="utf-8")


def abrir_navegador_normal():
    """Abre Edge/Chrome como lo abriría una persona y conecta Selenium por el puerto de depuración.
    Así el navegador no lleva las marcas del driver que Cloudflare detecta (error 600010)."""
    preparar_perfil()
    proceso = subprocess.Popen([
        buscar_navegador(),
        f"--remote-debugging-port={PUERTO_DEPURACION}",
        f"--user-data-dir={PERFIL_NAVEGADOR}",
        "--no-first-run", "--no-default-browser-check", "--start-maximized",
        "about:blank",
    ])
    # Espera a que el navegador abra el puerto
    fin = time.time() + 30
    while True:
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{PUERTO_DEPURACION}/json/version", timeout=2)
            break
        except Exception:
            if time.time() > fin:
                proceso.terminate()
                raise TimeoutError(f"{NOMBRE_NAVEGADOR} no abrió el puerto de depuración "
                                   f"{PUERTO_DEPURACION}; cierra las ventanas del RPA y reintenta")
            time.sleep(0.5)

    opciones, Driver = opciones_selenium()
    opciones.debugger_address = f"127.0.0.1:{PUERTO_DEPURACION}"
    driver = Driver(options=opciones)
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior",
                               {"behavior": "allow", "downloadPath": str(DESCARGAS)})
    except Exception:
        pass  # las preferencias del perfil ya fijan la carpeta de descargas
    driver.proceso_chrome = proceso  # para cerrarlo al final
    log(f"Paso 3 | {NOMBRE_NAVEGADOR} abierto (puerto {PUERTO_DEPURACION}) y conectado a Selenium")
    return driver


def cerrar_chrome(driver):
    driver.quit()
    proceso = getattr(driver, "proceso_chrome", None)
    if proceso:
        proceso.terminate()


def abrir_chrome_uc():
    """Chrome con SeleniumBase en modo UC (undetected-chromedriver)."""
    from seleniumbase import Driver as DriverUC  # se instala antes en instalar_seleniumbase()

    # external_pdf: descarga el PDF. eager: sigue apenas el HTML está listo (la DIAN tarda en
    # terminar de cargar y eso dejaba los clics esperando ~60 s)
    parametros = {"uc": True, "headless": False, "external_pdf": True, "page_load_strategy": "eager"}
    if NAVEGADOR_RUTA:
        parametros["binary_location"] = NAVEGADOR_RUTA
    driver = DriverUC(**parametros)
    driver.maximize_window()
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior",
                               {"behavior": "allow", "downloadPath": str(DESCARGAS)})
    except Exception:
        pass  # si no aplica, el PDF llega a downloaded_files/ y esperar_pdf lo trae a descargas/
    driver.modo_uc = True
    log("Paso 3 | Chrome abierto con SeleniumBase en modo UC")
    return driver


def instalar_seleniumbase():
    """Instala SeleniumBase en el mismo Python que ejecuta el RPA. Devuelve True si quedó listo.
    Prueba varias formas y guarda en rpa.log la salida de pip si todas fallan."""
    try:
        import seleniumbase  # noqa: F401
        return True
    except ModuleNotFoundError:
        pass
    log(f"Paso 3 | SeleniumBase no está instalado; lo instalo en {sys.executable}")
    py = sys.executable
    intentos = [
        [py, "-m", "ensurepip", "--upgrade"],  # por si este Python no trae pip
        [py, "-m", "pip", "install", "--upgrade", "pip"],
        [py, "-m", "pip", "install", "seleniumbase"],
        [py, "-m", "pip", "install", "--user", "seleniumbase"],  # sin permisos de administrador
        [py, "-m", "pip", "install", "--user", "--only-binary=:all:", "seleniumbase"],
    ]
    salida = ""
    for comando in intentos:
        r = subprocess.run(comando, capture_output=True, text=True)
        if "seleniumbase" in comando:
            salida = (r.stdout + r.stderr)[-3000:]
            if r.returncode == 0:
                break
    import importlib
    import site
    importlib.invalidate_caches()
    try:  # la carpeta --user puede no estar aún en sys.path
        sys.path.append(site.getusersitepackages())
    except Exception:
        pass
    try:
        import seleniumbase  # noqa: F401
        log("Paso 3 | SeleniumBase instalado")
        return True
    except ModuleNotFoundError:
        log("Paso 3 | No se pudo instalar SeleniumBase. Salida de pip:\n" + salida)
        return False


JS_POSICION_CASILLA = """
// Posición en pantalla de la casilla de Cloudflare (a ~30 px del borde izquierdo del widget)
const w = document.querySelector('[data-sitekey], .cf-turnstile, iframe[src*="challenges.cloudflare.com"]');
if (!w) return null;
w.scrollIntoView({block: 'center'});
const r = w.getBoundingClientRect();
if (r.width < 50) return null;
const borde = (window.outerWidth - window.innerWidth) / 2;
const barra = window.outerHeight - window.innerHeight - borde;
return {x: window.screenX + borde + r.left + 30, y: window.screenY + barra + r.top + r.height / 2,
        anchoPantalla: screen.width};
"""


def clic_con_mouse_real(driver):
    """Respaldo propio: clic con el mouse real (PyAutoGUI) sobre la casilla de Cloudflare."""
    try:
        import pyautogui
    except ModuleNotFoundError:
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", "pyautogui"],
                       capture_output=True)
        import site
        sys.path.append(site.getusersitepackages())
        import pyautogui
    pos = None
    for _ in range(12):  # espera a que el widget esté dibujado (máx. ~3 s)
        pos = driver.execute_script(JS_POSICION_CASILLA)
        if pos:
            break
        time.sleep(0.25)
    if not pos:
        raise RuntimeError("no encontré el widget de Cloudflare en la página")
    pyautogui.FAILSAFE = False  # con varios monitores el puntero puede pasar por una esquina
    escala = pyautogui.size()[0] / pos["anchoPantalla"]  # pantallas con zoom de Windows (125 %, 150 %)
    x, y = int(pos["x"] * escala), int(pos["y"] * escala)
    driver.switch_to.window(driver.current_window_handle)
    # Cloudflare falla ("La verificación falló") si detecta chromedriver conectado mientras
    # verifica: se desconecta el driver, se hace el clic y se reconecta después
    desconectar(driver)
    try:
        pyautogui.moveTo(x, y, duration=0.3)
        pyautogui.click()
        time.sleep(PAUSA_VERIFICACION)
    finally:
        conectar(driver)
    log(f"Paso 4 | Clic con el mouse real en la casilla ({x}, {y})")


def desconectar(driver):
    if es_uc(driver) and hasattr(driver, "disconnect"):
        try:
            driver.disconnect()
        except Exception:
            pass


def conectar(driver):
    if es_uc(driver) and hasattr(driver, "connect"):
        try:
            driver.connect()
        except Exception:
            pass


def pausa_desconectado(driver, segundos):
    """Deja el navegador sin chromedriver unos segundos (Cloudflare verifica sin detectarlo)."""
    if es_uc(driver) and hasattr(driver, "reconnect"):
        try:
            driver.reconnect(segundos)
            return
        except Exception:
            pass
    time.sleep(segundos)


def clic_captcha_uc(driver):
    """Hace clic en 'Verifique que es un ser humano'. Prueba los métodos de SeleniumBase que
    tenga la versión instalada y, si ninguno existe, usa el clic propio con el mouse real."""
    from seleniumbase.core import browser_launcher
    for nombre in ("uc_gui_click_captcha", "uc_gui_click_cf", "uc_gui_handle_captcha"):
        metodo = getattr(driver, nombre, None)
        if metodo is None and hasattr(browser_launcher, nombre):
            metodo = (lambda f: lambda: f(driver))(getattr(browser_launcher, nombre))
        if metodo:
            try:
                metodo()
                log(f"Paso 4 | Clic en la casilla con SeleniumBase ({nombre})")
                return
            except Exception as e:
                log(f"Paso 4 | {nombre} falló: {type(e).__name__}: {str(e)[:100]}")
    try:
        clic_con_mouse_real(driver)
    except Exception as e:
        log(f"Paso 4 | Clic con el mouse real falló: {type(e).__name__}: {str(e)[:100]}")


def funcion_uc(driver, nombre):
    """Devuelve un método UC de SeleniumBase. Algunas versiones no lo pegan al driver,
    así que se toma de seleniumbase.core.browser_launcher y se le pasa el driver."""
    metodo = getattr(driver, nombre, None)
    if metodo:
        return metodo
    from seleniumbase.core import browser_launcher
    funcion = getattr(browser_launcher, nombre)
    return lambda *args, **kwargs: funcion(driver, *args, **kwargs)


def es_uc(driver):
    return getattr(driver, "modo_uc", False)


def paso3_abrir_navegador(oculto=False):
    DESCARGAS.mkdir(exist_ok=True)
    if MODO_NAVEGADOR == "uc" and not oculto:
        if instalar_seleniumbase():
            return abrir_chrome_uc()
        log("Paso 3 | Sigo sin modo UC: abro Microsoft Edge normal conectado a Selenium")
        global NAVEGADOR, NOMBRE_NAVEGADOR, PERFIL_NAVEGADOR
        NAVEGADOR, NOMBRE_NAVEGADOR = "edge", "Microsoft Edge"
        PERFIL_NAVEGADOR = CARPETA / "perfil_edge"
        return abrir_navegador_normal()
    if MODO_NAVEGADOR == "normal" and not oculto:
        return abrir_navegador_normal()
    opciones, Driver = opciones_selenium()
    opciones.add_experimental_option("prefs", {
        "download.default_directory": str(DESCARGAS),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,  # descarga el PDF en vez de abrirlo
    })
    opciones.add_argument("--start-maximized")
    opciones.add_argument("--disable-blink-features=AutomationControlled")
    opciones.add_experimental_option("excludeSwitches", ["enable-automation"])
    if oculto:
        opciones.add_argument("--headless=new")
    return Driver(options=opciones)
