====================================================================================================
RPA DIAN – DESCARGA Y EXTRACCIÓN DE FACTURAS ELECTRÓNICAS
====================================================================================================

+--------------------------------------------------------------------------------------------------+
| NOTA DEL AUTOR                                                                                   |
|                                                                                                  |
| Mi mayor especialidad es el diseño de soluciones e innovación. Me especializo en herramientas    |
| de RPA como Automation Anywhere, UiPath y Power Automate, además de bases de datos SQL,          |
| Power Apps, Jira, ServiceNow, Azure, agentes de IA y documentación técnica.                      |
|                                                                                                  |
| Este proyecto es la primera vez que desarrollo un RPA con Python.                                |
+--------------------------------------------------------------------------------------------------+

Robot en Python que consulta facturas electrónicas por su CUFE en el
catálogo de la DIAN (https://catalogo-vpfe.dian.gov.co/User/SearchDocument), resuelve el
captcha de Cloudflare, descarga el PDF de la representación gráfica de cada factura, extrae sus
datos a un DataFrame por archivo y los exporta a CSV. Lleva el control en un Excel, permite
reprocesar los CUFEs que fallen, reporta tiempos, métricas y restricciones, y trae una consola de
operación y un dashboard de Power BI.

Resultado de referencia: 10 de 10 facturas descargadas al primer intento y 21 productos
extraídos, en 2 min 44 s (≈ 16 s por factura).


CONTENIDO
=========

1. Qué hace
2. Prerrequisitos
3. Instalación
4. Configuración (.env)
5. Uso
     - 5.1 Ejecutar el robot
     - 5.2 Reproceso de CUFEs fallidos
     - 5.3 Consola de operación
     - 5.4 Solo extracción de datos (punto 2)
     - 5.5 Dashboard de Power BI
6. Cómo funciona por dentro
7. Estructura del proyecto
8. Archivos que genera
9. Tiempos y métricas
10. Restricciones
11. Solución de problemas


1. QUÉ HACE
===========

lanzador.py ejecuta 7 pasos en orden. Cada paso está en su propio archivo:

  Paso  Archivo              Qué hace
  --------------------------------------------------------------------------------------------------
  1     paso1_correo.py      Lee el correo de la prueba (NIT del cuerpo y CUFEs del Word adjunto).
                             Si no hay correo configurado, usa cufes.txt y el NIT del .env.

  2     paso2_excel.py       Crea el Excel de control control_facturas.xlsx (hojas CUFES y NIT). Si
                             ya existe y tiene CUFEs marcados con X en REPROCESAR, lo conserva y
                             solo procesa esos.

  3     paso3_navegador.py   Abre Google Chrome con SeleniumBase en modo UC (no detectable por
                             Cloudflare).

  4     paso4_descarga.py    Consulta cada CUFE en la DIAN, resuelve el captcha, descarga el PDF y
                             marca Exitoso / No exitoso en el Excel. Hasta 5 intentos por CUFE.

  5     paso5_extraccion.py  Extrae el texto de cada factura (capa de texto u OCR), arma un
                             DataFrame de pandas por factura y lo exporta a CSV.

  6     paso6_reporte.py     Calcula tiempos y métricas, lista las restricciones y guarda el
                             reporte.

  7     paso7_envio.py       Envía el Excel de control por correo (opcional).


Variables extraídas de cada factura:

  Sección del PDF                     Variables
  --------------------------------------------------------------------------------------------------
  Datos del Documento / Datos del     numero_factura, fecha_emision, nit_emisor
  Emisor

  Detalles de Productos               codigo, descripcion, cantidad, precio_unitario


2. PRERREQUISITOS
=================

  Requisito                   Obligatorio  Detalle
  --------------------------------------------------------------------------------------------------
  Windows 10 u 11             Sí           Probado en Windows 11. El clic del captcha usa el mouse
                                           real del equipo.

  Python 3.10 o superior      Sí           Probado con Python 3.15. Debe estar en el PATH (python
                                           --version).

  Google Chrome               Sí           Versión actual. SeleniumBase descarga solo el driver que
                                           corresponde.

  Conexión a internet         Sí           Para el portal de la DIAN y para instalar librerías la
                                           primera vez.

  Pantalla y mouse libres     Sí           Mientras corre la descarga, no muevas el mouse ni cambies
                                           de ventana.

  Librerías de Python         Sí           selenium, seleniumbase, pandas, pdfplumber (ver
                                           requirements.txt). Si falta alguna, el robot la instala
                                           sola.

  Cuenta de Gmail con         No           Solo para leer el correo de la prueba (paso 1) y enviar
  contraseña de aplicación                 el Excel (paso 7). Crear contraseña de aplicación
                                           (https://myaccount.google.com/apppasswords).

  Tesseract OCR               No           Solo si llegan PDF escaneados (sin texto). En Windows:
                                           instalador UB Mannheim
                                           (https://github.com/UB-Mannheim/tesseract/wiki) con el
                                           idioma español.

  Power BI Desktop            No           Solo para el dashboard. Gratis en Microsoft Store.
  Microsoft Excel             No           Para abrir el Excel de control; debe estar cerrado
                                           mientras corre el robot.


3. INSTALACIÓN
==============

    # 1. Clonar el repositorio
    git clone https://github.com/JhonnathanDaza/Prueba-Outsoursing.git
    cd Prueba-Outsoursing
    
    # 2. (Opcional) entorno virtual
    python -m venv venv
    venv\Scripts\activate
    
    # 3. Librerías
    python -m pip install -r requirements.txt
    
    # 4. Configuración
    copy .env.example .env      # y edita .env (ver sección 4)

Si no vas a usar correo, deja los datos de correo vacíos en .env: el robot lee los CUFEs
de cufes.txt (uno por línea) y usa el NIT del .env.

Para verificar que todo quedó listo:

    python consola.py estado


4. CONFIGURACIÓN (.ENV)
=======================

Todas las variables son opcionales; los valores por defecto funcionan.

Datos de entrada y correo

  Variable          Por defecto    Para qué sirve
  --------------------------------------------------------------------------------------------------
  EMAIL_USUARIO     vacío          Cuenta de Gmail que lee el correo de la prueba y envía el Excel.
  EMAIL_CLAVE_APP   vacío          Contraseña de aplicación de Gmail (16 letras). Sin ella se usa
                                   cufes.txt y no se envía correo.

  CORREO_DESTINO    EMAIL_USUARIO  A quién se envía el Excel de control.
  CORREO_REMITENTE  vacío          Remitente del correo con los CUFEs. Vacío = cualquier correo con
                                   asunto "PRUEBA".

  NIT               900156264      NIT que se digita en la consulta. También es la contraseña de los
                                   PDF.


Navegador y captcha

  Variable            Por defecto  Para qué sirve
  --------------------------------------------------------------------------------------------------
  MODO_NAVEGADOR      uc           uc = Chrome con SeleniumBase en modo UC (recomendado). normal =
                                   Edge/Chrome normal conectado a Selenium. selenium = navegador
                                   lanzado por Selenium.

  NAVEGADOR           edge         Navegador para los modos normal y selenium (edge o chrome). El
                                   modo uc siempre usa Chrome.

  NAVEGADOR_RUTA      vacío        Ruta de chrome.exe / msedge.exe si no se encuentra sola.
  PUERTO_DEPURACION   9222         Puerto del modo normal.
  CASILLA_JS          si           Modos normal/selenium: marca la casilla con JavaScript antes de
                                   hacer clic.

  TWOCAPTCHA_API_KEY  vacío        Opcional: resuelve el captcha con el servicio de pago 2Captcha.


Tiempos y reintentos

  Variable               Por defecto  Para qué sirve
  --------------------------------------------------------------------------------------------------
  INTENTOS               5            Intentos por CUFE antes de marcarlo No exitoso (y con X para
                                      reprocesar).

  RECARGAS_CAPTCHA       5            Veces que recarga la página si la verificación no se completa.
  ESPERA_CAPTCHA         90           Segundos máximos de espera de la verificación (en modo UC se
                                      usan 25).

  RECONEXION_UC          2.5          Segundos que el driver queda desconectado mientras Cloudflare
                                      revisa la página. Súbelo a 4 si Cloudflare rechaza.

  PAUSA_VERIFICACION     3            Segundos sin driver después del clic en la casilla.
  PAUSA_ENTRE_CONSULTAS  0.5          Pausa entre facturas.


Extracción (punto 2)

  Variable        Por defecto  Para qué sirve
  --------------------------------------------------------------------------------------------------
  EXTRACCION      auto         auto = capa de texto y OCR si el PDF es imagen. texto = solo capa de
                               texto. ocr = siempre OCR.

  TESSERACT_RUTA  vacío        Ruta de tesseract.exe si no se encuentra sola.
  OCR_DPI         300          Resolución a la que se convierten las páginas para el OCR.


5. USO
======

Todos los comandos se ejecutan desde la carpeta del proyecto.


5.1 Ejecutar el robot
---------------------

    python lanzador.py               # proceso completo (o solo los CUFEs pendientes de reproceso)
    python lanzador.py --todo        # procesa todos los CUFEs aunque haya pendientes
    python lanzador.py --sin-correo  # no envía el correo al final

python rpa_dian.py también funciona (llama al lanzador). Mientras corre, no muevas el
mouse: el robot lo usa para marcar la casilla de Cloudflare. El avance se ve en la consola y
en rpa.log.


5.2 Reproceso de CUFEs fallidos
-------------------------------

El Excel de control tiene la columna REPROCESAR:

1. Si un CUFE falla después de todos sus intentos, queda No exitoso y con una X en
   REPROCESAR (en amarillo).
2. En la siguiente ejecución, si hay CUFEs con X, el robot procesa solo esos y conserva el
   resto de las filas.
3. Cuando se descarga bien, pasa a Exitoso y la X se borra.
4. Si no hay ninguna X, la ejecución procesa todos los CUFEs con un Excel nuevo.

Para forzar el reproceso de un CUFE, escribe X (mayúscula o minúscula) en su columna
REPROCESAR con el Excel cerrado, o usa python consola.py relanzar --cufe <N°>.


5.3 Consola de operación
------------------------

consola.py permite operar el robot sin abrir el código.

    python consola.py estado                 # valida si la plataforma está operativa
    python consola.py estado --json          # ... y el resultado en JSON
    
    python consola.py informe                # informe de la última ejecución
    python consola.py informe --guardar      # ... y lo guarda en salida/informe_ejecucion.txt y .json
    python consola.py informe --json         # ... en JSON
    
    python consola.py relanzar               # relanza los CUFEs marcados con X (o todo si no hay)
    python consola.py relanzar --cufe 4 7    # marca con X los CUFEs 4 y 7 y los relanza
    python consola.py relanzar --cufe c2e74  # también acepta el inicio del CUFE
    python consola.py relanzar --todo        # relanza todos los CUFEs desde cero
    python consola.py relanzar --sin-correo  # no envía el correo al final
    
    python consola.py                        # menú interactivo

- estado revisa: internet, portal de la DIAN (DNS, respuesta HTTP y tiempo), Chrome,
  librerías, Excel de control (que exista y no esté abierto), carpetas y espacio en disco,
  entrada de CUFEs, correo y si hay una ejecución en curso. Cada revisión sale como
  [ OK ], [AVISO] o [FALLA]. Que el portal salga [AVISO] porque Cloudflare pide
  verificación es normal: el robot la resuelve.
- informe muestra la última ejecución (inicio, fin, modo, intentos, errores, tiempo), los
  totales del Excel, el % de éxito, la extracción, el detalle por CUFE y los pendientes de
  reproceso.
- Códigos de salida: 0 = todo bien; 1 = hay fallas o CUFEs pendientes. Sirve para
  tareas programadas (por ejemplo, el Programador de tareas de Windows).


5.4 Solo extracción de datos (punto 2)
--------------------------------------

Para volver a extraer los datos de los PDF ya descargados, sin volver a consultar la DIAN:

    python paso5_extraccion.py             # capa de texto (OCR si algún PDF es imagen)
    python paso5_extraccion.py --ocr       # fuerza OCR sobre las imágenes de las páginas
    python paso5_extraccion.py --comparar  # exactitud del OCR frente a la capa de texto


5.5 Dashboard de Power BI
-------------------------

powerbi/DashboardFacturas.pbip lee control_facturas.xlsx y muestra:

- Tarjetas: total de CUFEs, exitosos, no exitosos, por reprocesar y % de éxito.
- Gráfica de torta con el porcentaje de éxito (Exitoso en verde, No exitoso en rojo,
  Pendiente en amarillo).
- Tabla con el detalle: N°, estado, fecha de proceso, REPROCESAR, observación y CUFE.

Uso: abre DashboardFacturas.pbip con Power BI Desktop y pulsa Actualizar después de cada
ejecución. Si mueves la carpeta del proyecto, cambia el parámetro RutaExcel en
Transformar datos → Administrar parámetros.


6. CÓMO FUNCIONA POR DENTRO
===========================


Captcha de Cloudflare (pasos 3 y 4)
-----------------------------------

La DIAN protege la consulta con Cloudflare Turnstile, que rechaza los navegadores
automatizados (error 600010). El robot lo resuelve así:

1. Abre Chrome con SeleniumBase en modo UC, un chromedriver modificado que oculta las
   señales de automatización.
2. Carga la página con el driver desconectado unos segundos (RECONEXION_UC) para que
   Cloudflare revise el navegador sin detectarlo. Normalmente la verificación pasa sola.
3. Si no pasa, ubica la casilla "Verifique que es un ser humano", desconecta el driver y
   hace clic con el mouse real (PyAutoGUI); luego se reconecta.
4. Si en 25 s no llega el token, recarga la página, vuelve a llenar el formulario y lo
   intenta otra vez (hasta RECARGAS_CAPTCHA veces).


Descarga del PDF (paso 4)
-------------------------

La página de la factura tiene su propia casilla de Cloudflare. Su botón Descargar PDF
muestra el aviso "Este archivo contiene contraseña…", y al pulsar Aceptar solo descarga
si la casilla ya tiene el token; si no, muestra "El campo de seguridad no está completo" y
reinicia la casilla. Por eso el robot:

1. Deja cargar la página con el driver desconectado.
2. Espera el token de Cloudflare (cf-turnstile-response); si no llega, marca la casilla.
3. Con el token listo, hace un clic en Descargar PDF y acepta el aviso de contraseña.
4. Espera el archivo; si la DIAN rechaza, espera un token nuevo y repite.


Extracción de datos (paso 5)
----------------------------

1. Abre el PDF con su contraseña (el NIT consultado).
2. Capa de texto (por defecto): pdfplumber lee el texto y la tabla con la posición de
   cada carácter. El PDF corta las descripciones a ancho fijo, pero conserva el espacio al
   final de la línea cuando el corte es entre palabras, así que la descripción se reconstruye
   exacta.
3. OCR (PDF escaneados o --ocr): cada página se convierte en imagen a 300 dpi, se limpia
   (se quitan las líneas y fondos verdes de la tabla) y se lee con Tesseract. Las columnas se
   ubican con las líneas verticales de la tabla y cada fila se ancla en su número (Nro.).
4. Encabezado con expresiones regulares (número de factura, fecha de emisión, NIT del emisor).
   La tabla de productos se reconoce por los títulos de sus columnas.
5. Normaliza: fecha a ISO (AAAA-MM-DD) y valores a número ($ 2.900.000,00 → 2900000.0).
6. Un DataFrame de pandas por factura → salida/csv_por_factura/<numero_factura>.csv, más
   el consolidado.


7. ESTRUCTURA DEL PROYECTO
==========================

    Prueba-Outsoursing/
    ├── lanzador.py               Ejecuta los pasos 1 a 7 en orden y mide los tiempos
    ├── consola.py                Relanzamientos, informe de ejecución y estado de la plataforma
    ├── config.py                 Configuración compartida: .env, rutas, parámetros, log()
    ├── paso1_correo.py           1. Leer el correo de la prueba (NIT + CUFEs)
    ├── paso2_excel.py            2. Excel de control y lógica de reproceso
    ├── paso3_navegador.py        3. Abrir Chrome (modo UC) y utilidades del captcha
    ├── paso4_descarga.py         4. Consultar cada CUFE, resolver el captcha y descargar el PDF
    ├── paso5_extraccion.py       5. Extraer texto (texto/OCR), DataFrame por factura y CSV
    ├── paso6_reporte.py          6. Tiempos, métricas y restricciones
    ├── paso7_envio.py            7. Enviar el Excel por correo
    ├── rpa_dian.py               Compatibilidad: llama al lanzador
    ├── cufes.txt                 CUFEs de respaldo (si no se lee el correo)
    ├── requirements.txt          Librerías de Python
    ├── .env.example              Plantilla de configuración (copiar como .env)
    ├── powerbi/                  Dashboard de Power BI (DashboardFacturas.pbip)
    └── README.md


8. ARCHIVOS QUE GENERA
======================

    control_facturas.xlsx                  Excel de control: estado de cada CUFE y columna REPROCESAR
    descargas/<cufe>.pdf                   PDF descargados
    descargas/error_<n>_<intento>.png      Captura de pantalla cuando falla un intento
    salida/csv_por_factura/<numero>.csv    Un CSV por factura (un DataFrame por archivo)
    salida/facturas_consolidado.csv        Todas las facturas juntas
    salida/metricas_descarga.csv           Estado, intentos y tiempos por CUFE
    salida/metricas_extraccion.csv         Método, productos, campos completos y tiempos por PDF
    salida/reporte.txt                     Tiempos, métricas y restricciones de la ejecución
    salida/informe_ejecucion.txt / .json   Informe generado con consola.py informe --guardar
    salida/comparacion_texto_vs_ocr.csv    Exactitud del OCR (paso5_extraccion.py --comparar)
    rpa.log                                Bitácora de todas las ejecuciones (la última al final)


9. TIEMPOS Y MÉTRICAS
=====================

Ejecución completa (10 CUFEs):

  Métrica               Valor
  --------------------------------------------------------------------------------------------------
  Facturas descargadas  10 de 10 (100 %), 1 intento cada una
  Tiempo total          2 min 44 s
  Promedio por factura  ≈ 16 s (de ellos ≈ 13 s en la verificación de Cloudflare)
  Extracción de datos   10 de 10 completas, 21 productos, ≈ 0,37 s por PDF
  Reproceso de 1 CUFE   ≈ 25 s


Extracción: capa de texto frente a OCR (10 facturas, 21 productos):

                              Capa de texto  OCR (Tesseract, 300 dpi)
  --------------------------------------------------------------------------------------------------
  Tiempo por PDF              0,37 s         3,6 s
  Número de factura / fecha   100 %          100 %
  / NIT

  Precio unitario             100 %          100 %
  Cantidad                    100 %          95,2 %
  Código                      100 %          66,7 %
  Descripción                 100 %          87,9 % de similitud (47,6 % exacta)


Métricas que se reportan en cada ejecución (salida/reporte.txt): facturas consultadas,
exitosas y no exitosas, tasa de éxito, intentos, PDFs con datos completos, productos
extraídos, método de extracción, productos sin código, tiempo por etapa, tiempo en captcha,
promedio por factura y tiempos de extracción (promedio, mínimo y máximo).


10. RESTRICCIONES
=================

- Cloudflare: el captcha se resuelve con SeleniumBase en modo UC y el mouse real, así que
  el equipo no se puede usar mientras corre la descarga y no funciona en modo oculto
  (headless). Cloudflare puede cambiar su detección en cualquier momento.
- HTML de la DIAN: los selectores (formulario, Descargar PDF, avisos) dependen del sitio
  actual; si cambia, hay que ajustarlos en paso4_descarga.py.
- PDF con contraseña: se abren con el NIT del .env (el NIT digitado en la consulta).
- Extracción: la tabla de productos se reconoce por los títulos de sus columnas; si cambia
  el formato de la representación gráfica, hay que ajustar columnas_de().
- OCR: solo para PDF escaneados; requiere Tesseract y confunde caracteres parecidos
  (S/5, 8/B, O/0) y palabras cortadas a mitad de línea.
- Datos de origen: una factura (HTFE-6599) no trae código en su producto; queda vacío y se
  reporta como producto sin código, no como error.
- Excel de control: debe estar cerrado mientras corre el robot.
- Ritmo: 0,5 s entre consultas y hasta 5 intentos por CUFE, para no saturar el portal.


11. SOLUCIÓN DE PROBLEMAS
=========================

  Síntoma                             Causa y solución
  --------------------------------------------------------------------------------------------------
  ModuleNotFoundError: seleniumbase   El robot la instala solo. Si falla, ejecuta python -m pip
  (u otra librería)                   install -r requirements.txt con el mismo Python que corre el
                                      robot.

  Error 600010 / "La verificación     Cloudflare detectó el navegador. Usa MODO_NAVEGADOR=uc, no
  falló"                              muevas el mouse y sube RECONEXION_UC a 4.

  "El campo de seguridad no está      La casilla de la página de la factura no se había completado.
  completo"                           El robot espera el token y reintenta solo.

  "No encontré Google Chrome"         Pon la ruta de chrome.exe en NAVEGADOR_RUTA.
  "El archivo está abierto en Excel"  Cierra control_facturas.xlsx.
  Un CUFE queda No exitoso            Queda con X en REPROCESAR: ejecuta python consola.py relanzar.
                                      Revisa descargas/error_<n>_<intento>.png y rpa.log.

  "no tiene capa de texto y no está   El PDF es una imagen: instala Tesseract y define
  instalado Tesseract"                TESSERACT_RUTA.

  El dashboard sale vacío o con       Pulsa Actualizar; si moviste la carpeta, cambia el parámetro
  error de ruta                       RutaExcel.

  ¿Está todo listo para correr?       python consola.py estado
