"""
toques_y_selenium.py
1) Conecta la VPN de Ivanti (iconos ocultos -> Abrir -> Conectar -> contraseña).
2) Espera a que TÚ ingreses el código del autenticador.
3) Abre el portal de distribuidores en Edge y hace clic en la opción SIAC.

Uso:
  python toques_y_selenium.py                   # flujo completo: VPN -> token -> portal -> SIAC
  python toques_y_selenium.py --solo-selenium   # VPN ya conectada: solo portal -> SIAC
  python toques_y_selenium.py --guardar-password  # 1ra vez: guarda la contraseña de la VPN de forma segura
  python toques_y_selenium.py --listar          # solo lista las apps de iconos ocultos (sin Selenium)
"""
import argparse
import base64
import csv
import logging
import re
import sys
import time
from pathlib import Path

import keyring
import psutil
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ============================ CONFIG ============================
# Texto del tooltip del icono que quieres clicar (usa --listar para ver los nombres exactos)
REGEX_ICONO = re.compile(r"Ivanti Secure Access", re.I)
TIPO_CLIC = "left"            # "left", "right" o "double"
# Opción del menú que aparece después del clic en el icono. None = no clicar nada
REGEX_MENU = re.compile(r"(Abrir|Open) Ivanti Secure Access", re.I)
ESPERA_ANTES_SELENIUM = 5     # segundos a esperar después del clic (ej. que conecte la VPN)

# Ventana de Ivanti que se abre tras el menú
REGEX_VENTANA_IVANTI = re.compile(r"Ivanti Secure Access", re.I)
REGEX_BOTON_CONECTAR = re.compile(r"^\s*(Conectar|Connect)\s*$", re.I)       # exacto: no confunde con "Desconectar"
REGEX_BOTON_DESCONECTAR = re.compile(r"^\s*(Desconectar|Disconnect)\s*$", re.I)
# Login de la VPN: la contraseña se guarda en el Administrador de credenciales de Windows
# (python toques_y_selenium.py --guardar-password), NUNCA en este archivo.
USUARIO_VPN = "D79983538"
KEYRING_SERVICIO = "ivanti_vpn"
REGEX_VENTANA_LOGIN = re.compile(r"(Ivanti Secure Access|Conect[aá]ndose a|Connecting to)", re.I)
TIMEOUT_LOGIN = 15            # segundos esperando el diálogo de usuario/contraseña
PAUSA_TECLA = 0.15            # segundos entre tecla y tecla (súbelo si el diálogo se pierde caracteres)
ESPERA_FOCO = 3               # segundos antes de empezar a teclear (que el diálogo tome el foco)
# 2do factor: el código del autenticador lo ingresas TÚ (no se automatiza, es el punto del MFA)
REGEX_DIALOGO_TOKEN = re.compile(r"(token|credenciales para completar|secundari)", re.I)
TIMEOUT_TOKEN = 180           # segundos para que ingreses el código del autenticador
INDICE_CONEXION = 0           # si tienes varias conexiones con botón "Conectar", cuál usar (0 = la primera)
TIMEOUT_VENTANA = 20          # segundos máximos esperando que aparezca la ventana
TIMEOUT_CONEXION = 120        # segundos máximos esperando que la VPN quede conectada (incluye login/MFA manual)
# Respaldo si Ivanti no expone el botón: coordenadas del botón RELATIVAS a la ventana de Ivanti.
# Obténlas con: python toques_y_selenium.py --capturar-ivanti   (con la ventana de Ivanti abierta)
COORDS_CONECTAR = (256, 134)  # estimado de tu captura (botón Conectar de la 1ra conexión); verifícalo con --capturar-ivanti

# Nombre del botón ^ (tooltip o nombre interno, según idioma/versión de Windows)
REGEX_CHEVRON = re.compile(r"(Mostrar [ií]conos ocultos|Show hidden icons|Notification Chevron|Chevron)", re.I)
# Clase de la ventana de iconos ocultos: Windows 10 / Windows 11
CLASES_OVERFLOW = ("NotifyIconOverflowWindow", "TopLevelWindowForOverflowXamlIsland")

URL_SELENIUM = "http://portaldistribuidores.tim.com.pe:1616/SitePages/inicio.aspx"
TEXTO_OPCION = "SIAC Unico"   # opción a buscar y clicar dentro del portal
TIPO_BUSQUEDA = "Documento Identidad"   # opción del desplegable de búsqueda en SIAC Único
VALOR_BUSQUEDA = ""           # un solo valor. Vacío = itera el Excel de abajo
# Excel con los valores a buscar (en la misma carpeta que este script)
ARCHIVO_VALORES = "RUCS_PORTA_20MAS_100.xlsx"
COLUMNA_VALORES = None        # None = autodetecta (RUC/documento/DNI); o el nombre o índice de la columna
HOJA_VALORES = 0              # índice o nombre de la hoja
PAUSA_ENTRE_BUSQUEDAS = 3     # segundos entre una búsqueda y la siguiente
# Popup "Consulta de Productos" que sale al presionar la lupa
TIMEOUT_POPUP = 60            # segundos esperando que aparezca el popup
REGEX_POPUP = re.compile(r"popup|consulta de productos", re.I)
TIMEOUT_DATOS = 30            # segundos esperando que el popup termine de pintar sus tablas
ESPERA_DATOS = 2              # segundos entre lectura y lectura del popup
CLIC_EN_TOTAL_ACTIVAS = True  # tras leer los totales, clic en el número de activas (abre otro popup)
ARCHIVO_CSV = "resultado_rucs.csv"
# El total se repite en cada producto de la misma cuenta: True lo cuenta una sola vez
AGRUPAR_POR_CUENTA = True
TIMEOUT_WEB = 40              # segundos esperando la carga del portal / la opción
TIMEOUT_VENTANA_SIAC = 30     # segundos esperando que SIAC abra su ventana nueva
REGEX_VENTANA_SIAC = re.compile(r"siac", re.I)   # título o URL de la ventana de SIAC
# Dominios donde Edge puede usar tu sesión de Windows (evita el popup de usuario/contraseña)
AUTH_ALLOWLIST = "*tim.com.pe,portaldistribuidores.tim.com.pe"
# Credenciales del portal: por defecto reutiliza las mismas de la VPN
REUSAR_CREDENCIALES_VPN = True        # False si el portal usa otro usuario/contraseña
USUARIO_PORTAL = USUARIO_VPN          # ej. "TIM\\D79983538" si el portal pide dominio
KEYRING_SERVICIO_PORTAL = "portal_distribuidores"
REGEX_DIALOGO_WEB = re.compile(r"(Seguridad de Windows|Windows Security|Iniciar sesi[oó]n|Sign in)", re.I)
NAVEGADOR = "chrome"          # "chrome" o "edge"
# Perfil dedicado para conservar cookies/sesión (SSO) entre ejecuciones. None = perfil temporal
PERFIL_DIR = Path(f"./perfil_{NAVEGADOR}").resolve()
MANTENER_ABIERTO = True       # deja el navegador abierto al terminar el script
# ================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("bandeja")


# ---------------------- bandeja de Windows ----------------------
def _escritorio():
    return Desktop(backend="uia")


def _buscar_boton(contenedor, regex):
    for boton in contenedor.descendants(control_type="Button"):
        if regex.search(boton.window_text() or ""):
            return boton
    return None


def _ventana_overflow_visible():
    for w in _escritorio().windows():
        try:
            if w.class_name() in CLASES_OVERFLOW and w.is_visible():
                return w
        except Exception:
            continue
    return None


def _buscar_chevron(barra):
    # 1) Por nombre
    chevron = _buscar_boton(barra, REGEX_CHEVRON)
    if chevron:
        return chevron
    # 2) Por estructura (Windows 10): el ^ es el control de clase "Button" dentro de TrayNotifyWnd
    try:
        area = barra.child_window(class_name="TrayNotifyWnd").wrapper_object()
        for b in area.descendants():
            if b.class_name() == "Button":
                return b
    except Exception as e:
        log.debug("Búsqueda por estructura falló: %s", e)
    return None


def abrir_iconos_ocultos(timeout=5):
    """Clic en el botón ^ de la barra de tareas y devuelve la ventana de iconos ocultos."""
    ya_abierta = _ventana_overflow_visible()
    if ya_abierta:
        return ya_abierta

    barra = _escritorio().window(class_name="Shell_TrayWnd")
    chevron = _buscar_chevron(barra)
    if chevron is None:
        nombres = [f"{b.window_text()!r} [{b.class_name()}]" for b in barra.descendants(control_type="Button")]
        raise RuntimeError(
            "No encontré el botón 'Mostrar iconos ocultos'. ¿Está activada 'Mostrar siempre todos los iconos "
            "en el área de notificación'? Botones de la barra: " + ", ".join(nombres)
        )
    log.info("[1/8] Bandeja: abriendo iconos ocultos (%s)", chevron.window_text())
    chevron.click_input()

    fin = time.time() + timeout
    while time.time() < fin:
        ventana = _ventana_overflow_visible()
        if ventana:
            return ventana
        time.sleep(0.3)
    raise RuntimeError("No apareció la ventana de iconos ocultos.")


def listar_iconos_ocultos():
    """Abre los iconos ocultos, imprime las apps encontradas y devuelve la lista de nombres."""
    ventana = abrir_iconos_ocultos()
    apps = []
    for boton in ventana.descendants(control_type="Button"):
        texto = (boton.window_text() or "").strip()
        if texto and texto not in apps:
            apps.append(texto)
    send_keys("{ESC}")  # cierra la ventana de iconos ocultos

    if not apps:
        log.warning("No se encontraron apps en iconos ocultos.")
        return apps

    print(f"\nApps en iconos ocultos ({len(apps)}):")
    for i, texto in enumerate(apps, 1):
        # El tooltip suele traer "Nombre\nEstado": se muestra el nombre y el estado entre paréntesis
        lineas = [l.strip() for l in texto.splitlines() if l.strip()]
        detalle = f"  ({' | '.join(lineas[1:])})" if len(lineas) > 1 else ""
        print(f"  {i:>2}. {lineas[0]}{detalle}")
    print()
    return apps


def clic_menu_contextual(regex, timeout=5):
    """Busca y clica la opción en el menú que se abre tras el clic en el icono."""
    vistos = set()
    fin = time.time() + timeout
    while time.time() < fin:
        for w in _escritorio().windows():
            try:
                if w.class_name() != "#32768" and w.element_info.control_type != "Menu":
                    continue
                for item in w.descendants(control_type="MenuItem"):
                    texto = item.window_text() or ""
                    vistos.add(texto)
                    if regex.search(texto):
                        log.info("[2/8] Ivanti: opción del menú -> %s", texto)
                        item.click_input()
                        return
            except Exception:
                continue
        time.sleep(0.3)
    send_keys("{ESC}")
    raise RuntimeError(
        f"No encontré la opción '{regex.pattern}' en el menú. "
        f"Opciones vistas: {', '.join(sorted(t for t in vistos if t)) or 'ninguna (el menú no se detectó)'}"
    )


def clic_icono_oculto():
    ventana = abrir_iconos_ocultos()
    icono = _buscar_boton(ventana, REGEX_ICONO)
    if icono is None:
        disponibles = [b.window_text().splitlines()[0] for b in ventana.descendants(control_type="Button")
                       if b.window_text()]
        send_keys("{ESC}")
        raise RuntimeError(
            f"No encontré '{REGEX_ICONO.pattern}' en iconos ocultos. "
            f"Apps disponibles: {', '.join(disponibles) or 'ninguna'}. "
            "¿La app está abierta y su icono está oculto (no fijado en la barra)?"
        )

    log.info("[2/8] Ivanti: clic en el icono %s", icono.window_text())
    if TIPO_CLIC == "double":
        icono.double_click_input()
    elif TIPO_CLIC == "right":
        icono.right_click_input()
    else:
        icono.click_input()

    if REGEX_MENU:
        clic_menu_contextual(REGEX_MENU)


# ----------------------- ventana de Ivanti -----------------------
TIPOS_BOTON = ("Button", "Hyperlink", "SplitButton")


def _botones(ventana, regex):
    """Controles cuyo texto coincide: primero botones/enlaces; si no hay, cualquier control (texto, celda...)."""
    todos = []
    for el in ventana.descendants():
        try:
            if regex.search(el.window_text() or ""):
                todos.append(el)
        except Exception:
            continue
    botones = [el for el in todos if el.element_info.control_type in TIPOS_BOTON]
    return botones or todos


def _interfaces_activas():
    """Nombres de adaptadores de red que están arriba (la VPN crea/levanta uno al conectar)."""
    return {nombre for nombre, st in psutil.net_if_stats().items() if st.isup}


def esperar_ventana_ivanti(timeout=TIMEOUT_VENTANA):
    fin = time.time() + timeout
    while time.time() < fin:
        for w in _escritorio().windows():
            try:
                if w.is_visible() and REGEX_VENTANA_IVANTI.search(w.window_text() or ""):
                    return w
            except Exception:
                continue
        time.sleep(0.5)
    raise RuntimeError("No apareció la ventana de Ivanti Secure Access Client.")


def clic_conectar_ivanti():
    ventana = esperar_ventana_ivanti()
    log.info("[3/8] Ivanti: ventana abierta (%s)", ventana.window_text())
    ventana.set_focus()
    time.sleep(1)  # que termine de pintar la lista de conexiones

    interfaces_antes = _interfaces_activas()
    botones = _botones(ventana, REGEX_BOTON_CONECTAR)
    if botones:
        if len(botones) > 1:
            log.info("Hay %d botones Conectar; usando el índice %d (INDICE_CONEXION).", len(botones), INDICE_CONEXION)
        log.info("[3/8] Ivanti: clic en %s [%s]", botones[INDICE_CONEXION].window_text(),
                 botones[INDICE_CONEXION].element_info.control_type)
        botones[INDICE_CONEXION].click_input()
    elif _botones(ventana, REGEX_BOTON_DESCONECTAR):
        log.info("[3/8] VPN ya conectada; omito el clic en Conectar.")
        return
    elif COORDS_CONECTAR:
        log.info("No detecté el botón por nombre; clic en coordenadas %s de la ventana.", COORDS_CONECTAR)
        ventana.click_input(coords=COORDS_CONECTAR)
    else:
        controles = [f"{el.element_info.control_type}:{el.window_text()!r}"
                     for el in ventana.descendants() if el.window_text()][:40]
        raise RuntimeError(
            "No encontré el botón Conectar. Controles con texto en la ventana: "
            + (", ".join(controles) or "ninguno (Ivanti no expone sus controles; usa COORDS_CONECTAR)")
        )
    ingresar_password()
    esperar_token_mfa()
    log.info("[5/8] Esperando que la VPN quede activa (hasta %s s)...", TIMEOUT_CONEXION)

    fin = time.time() + TIMEOUT_CONEXION
    while time.time() < fin:
        nuevas = _interfaces_activas() - interfaces_antes
        if nuevas:
            log.info("[5/8] VPN CONECTADA (adaptador: %s).", ", ".join(sorted(nuevas)))
            return
        try:
            if _botones(ventana, REGEX_BOTON_DESCONECTAR):
                log.info("[5/8] VPN CONECTADA.")
                return
        except Exception:
            pass  # la ventana puede refrescarse o cerrarse mientras conecta
        time.sleep(2)
    log.warning("[5/8] No confirmé la VPN en %s s; continúo igual.", TIMEOUT_CONEXION)


def _es_password(el):
    try:
        return bool(el.element_info.element.CurrentIsPassword)
    except Exception:
        return False


def _escapar_teclas(texto):
    """Escapa los caracteres especiales de type_keys (+ ^ % ~ ( ) { } [ ])."""
    return "".join("{%s}" % c if c in "+^%~(){}[]" else c for c in texto)


def esperar_dialogo_login(timeout=TIMEOUT_LOGIN):
    """Devuelve (ventana, campo_contraseña) del diálogo de login de Ivanti, o (None, None)."""
    fin = time.time() + timeout
    while time.time() < fin:
        for w in _escritorio().windows():
            try:
                if not (w.is_visible() and REGEX_VENTANA_LOGIN.search(w.window_text() or "")):
                    continue
                edits = w.descendants(control_type="Edit")
                campo = next((e for e in edits if _es_password(e)), None)
                if campo is None and len(edits) >= 2:
                    campo = edits[-1]  # usuario y contraseña: la contraseña es el último campo
                if campo is not None:
                    return w, campo
            except Exception:
                continue
        time.sleep(0.5)
    return None, None


def ingresar_password():
    ventana, campo = esperar_dialogo_login()
    if campo is None:
        log.info("[4/8] Ivanti no pidió contraseña; continúo.")
        return

    password = keyring.get_password(KEYRING_SERVICIO, USUARIO_VPN)
    if not password:
        raise RuntimeError(
            f"No hay contraseña guardada para {USUARIO_VPN}. "
            "Ejecuta primero: python toques_y_selenium.py --guardar-password"
        )

    log.info("[4/8] Ivanti: ingresando contraseña de la VPN...")
    campo.click_input()
    campo.type_keys("^a{BACKSPACE}", pause=0.02)            # limpia el campo por si tenía algo
    campo.type_keys(_escapar_teclas(password), with_spaces=True, pause=0.02)
    del password

    botones = _botones(ventana, REGEX_BOTON_CONECTAR)
    if botones:
        botones[0].click_input()
    else:
        send_keys("{ENTER}")
    log.info("[4/8] Ivanti: contraseña enviada.")


def esperar_token_mfa(timeout=TIMEOUT_TOKEN):
    """Si Ivanti pide el código del autenticador, espera a que TÚ lo ingreses."""
    fin_deteccion = time.time() + TIMEOUT_LOGIN
    dialogo = None
    while time.time() < fin_deteccion and dialogo is None:
        for w in _escritorio().windows():
            try:
                if not (w.is_visible() and REGEX_VENTANA_LOGIN.search(w.window_text() or "")):
                    continue
                textos = " ".join(el.window_text() or "" for el in w.descendants())
                if REGEX_DIALOGO_TOKEN.search(textos):
                    dialogo = w
                    break
            except Exception:
                continue
        if dialogo is None:
            time.sleep(0.5)

    if dialogo is None:
        return  # no pidió segundo factor

    log.info("[5/8] >> TU TURNO: ingresa el código del autenticador y dale Conectar.")
    log.info(">> Esperando hasta %s s...", TIMEOUT_TOKEN)
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            if not dialogo.is_visible():   # el diálogo se cierra al aceptar el código
                log.info("[5/8] Código ingresado; continúo.")
                return
        except Exception:
            log.info("[5/8] Código ingresado; continúo.")
            return
        time.sleep(1)
    log.warning("[5/8] Se agotó el tiempo esperando el código; continúo igual.")


def guardar_password():
    import getpass
    pwd = getpass.getpass(f"Contraseña de la VPN para {USUARIO_VPN} (no se mostrará): ")
    if not pwd:
        print("Contraseña vacía; no se guardó nada.")
        return
    keyring.set_password(KEYRING_SERVICIO, USUARIO_VPN, pwd)
    print("Guardada en el Administrador de credenciales de Windows.")


def capturar_coords_ivanti():
    """Muestra en vivo la posición del mouse relativa a la ventana de Ivanti. Ctrl+C para salir."""
    import win32api  # viene con pywin32 (dependencia de pywinauto)
    ventana = esperar_ventana_ivanti(timeout=5)
    print("Pon el mouse sobre el botón Conectar (sin hacer clic). Ctrl+C para terminar.\n")
    try:
        while True:
            r = ventana.rectangle()
            x, y = win32api.GetCursorPos()
            print(f"\rCOORDS_CONECTAR = ({x - r.left}, {y - r.top})      ", end="", flush=True)
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nCopia el último valor en COORDS_CONECTAR (sección CONFIG).")


# -------------------------- Selenium --------------------------
def abrir_selenium(url=URL_SELENIUM, navegar=True):
    if NAVEGADOR == "edge":
        opts = webdriver.EdgeOptions()
    else:
        opts = webdriver.ChromeOptions()

    opts.add_argument("--start-maximized")
    # 'none': get() no espera la carga, así podemos atender el popup de autenticación
    opts.set_capability("pageLoadStrategy", "none")
    if AUTH_ALLOWLIST:
        # Autenticación integrada de Windows: Edge responde solo al reto NTLM/Kerberos
        opts.add_argument(f"--auth-server-allowlist={AUTH_ALLOWLIST}")
        opts.add_argument(f"--auth-negotiate-delegate-allowlist={AUTH_ALLOWLIST}")
    if PERFIL_DIR:
        opts.add_argument(f"--user-data-dir={PERFIL_DIR}")
    if MANTENER_ABIERTO:
        opts.add_experimental_option("detach", True)

    # Selenium 4.6+ descarga el driver solo (Selenium Manager)
    driver = webdriver.Edge(options=opts) if NAVEGADOR == "edge" else webdriver.Chrome(options=opts)
    if navegar:
        driver.get(url)
        WebDriverWait(driver, TIMEOUT_WEB).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        log.info("Selenium abierto en: %s", driver.current_url)
    return driver


def _buscar_en_frames(driver, xpath, timeout=TIMEOUT_WEB):
    """Busca el xpath en la página y dentro de cada iframe (SharePoint suele usarlos)."""
    fin = time.time() + timeout
    while time.time() < fin:
        driver.switch_to.default_content()
        contextos = [None] + driver.find_elements(By.TAG_NAME, "iframe")
        for frame in contextos:
            try:
                driver.switch_to.default_content()
                if frame is not None:
                    driver.switch_to.frame(frame)
                elementos = [el for el in driver.find_elements(By.XPATH, xpath) if el.is_displayed()]
                if elementos:
                    return elementos[0]
            except Exception:
                continue
        time.sleep(1)
    driver.switch_to.default_content()
    return None


_MAYUS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ"
_MINUS = "abcdefghijklmnopqrstuvwxyzaeiouun"


def _normaliza(texto):
    return texto.lower().translate(str.maketrans("áéíóúüñ", "aeiouun"))


def _xpath_texto(tags, texto):
    """XPath que busca 'texto' en esos tags, sin distinguir mayúsculas ni tildes."""
    obj = _normaliza(texto)
    tr = f"translate(normalize-space(.),'{_MAYUS}','{_MINUS}')"
    return " | ".join(f"//{t}[contains({tr},'{obj}')]" for t in tags)


# Textos posibles del botón del desplegable (el tipo de búsqueda activo)
OPCIONES_BUSQUEDA = ("Teléfono", "Telefono", "Cuenta", "Customer Id", "ICCID", "Nombres",
                     "Documento Identidad", "Número de Recibo", "Razón Social")


def _buscar_toggle(driver):
    """El botón del desplegable de tipo de búsqueda (no el menú del usuario)."""
    # 1) Por su texto actual: es el único que muestra un tipo de búsqueda
    for opcion in OPCIONES_BUSQUEDA:
        xp = _xpath_texto(("button", "a", "span", "div"), opcion)
        for el in driver.find_elements(By.XPATH, xp):
            try:
                if el.is_displayed() and len(el.text.strip()) < 25:
                    return el
            except Exception:
                continue
    # 2) Respaldo: un dropdown-toggle visible con texto corto
    for el in driver.find_elements(By.XPATH, "//*[contains(@class,'dropdown-toggle')]"
                                             " | //*[@data-toggle='dropdown']"):
        try:
            if el.is_displayed() and 0 < len(el.text.strip()) < 25:
                return el
        except Exception:
            continue
    return None


def seleccionar_tipo_busqueda(driver, texto=TIPO_BUSQUEDA):
    """En SIAC Único: abre el desplegable de tipo de búsqueda y elige la opción indicada."""
    log.info("[8/8] SIAC: buscando el botón del desplegable...")
    toggle = _buscar_toggle(driver)
    if toggle is None:
        raise RuntimeError("No encontré el desplegable de tipo de búsqueda en SIAC.")

    actual = toggle.text.strip()
    log.info("[8/8] SIAC: botón encontrado -> %r. Haciendo clic para desplegarlo...", actual)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", toggle)
    time.sleep(0.5)
    try:
        toggle.click()
    except Exception:
        driver.execute_script("arguments[0].click();", toggle)
    time.sleep(1.5)
    log.info("[8/8] SIAC: desplegable abierto. Buscando la opción %r...", texto)

    # La opción solo existe/es visible con el menú abierto
    opcion = None
    for el in driver.find_elements(By.XPATH, _xpath_texto(("a", "li", "span", "button"), texto)):
        try:
            if el.is_displayed():
                opcion = el
                break
        except Exception:
            continue

    if opcion is None:
        visibles = sorted({el.text.strip() for el in driver.find_elements(By.XPATH, "//li | //a")
                           if el.text.strip() and len(el.text.strip()) < 40 and el.is_displayed()})[:30]
        raise RuntimeError(
            f"Abrí el desplegable pero no encontré '{texto}'. Opciones visibles: {', '.join(visibles)}"
        )

    log.info("[8/8] SIAC: opción encontrada -> %r. Haciendo clic...", opcion.text.strip())
    try:
        opcion.click()
    except Exception:
        driver.execute_script("arguments[0].click();", opcion)
    time.sleep(1)
    nuevo = _buscar_toggle(driver).text.strip()
    if _normaliza(texto) in _normaliza(nuevo):
        log.info("[8/8] LISTO: el tipo de búsqueda cambió de %r a %r", actual, nuevo)
    else:
        log.warning("[8/8] El botón sigue mostrando %r; puede que el clic no aplicara.", nuevo)
    return opcion


def buscar_en_siac(driver, valor=VALOR_BUSQUEDA):
    """Escribe el valor en la caja de búsqueda de SIAC y hace clic en la lupa."""
    if not valor:
        log.info("[8/8] SIAC: sin valor que buscar (VALOR_BUSQUEDA está vacío).")
        return False

    # La caja de texto de la barra roja (junto al desplegable)
    caja = None
    for el in driver.find_elements(By.XPATH, "//input[@type='text'] | //input[not(@type)]"):
        try:
            if el.is_displayed() and el.is_enabled():
                caja = el
                break
        except Exception:
            continue
    if caja is None:
        raise RuntimeError("No encontré la caja de búsqueda en SIAC.")

    log.info("[8/8] SIAC: caja de búsqueda encontrada (name=%r, id=%r).",
             caja.get_attribute("name"), caja.get_attribute("id"))
    log.info("[8/8] SIAC: escribiendo %r en el input...", valor)
    caja.clear()
    caja.send_keys(valor)
    time.sleep(0.5)

    escrito = caja.get_attribute("value")
    if escrito == valor:
        log.info("[8/8] SIAC: dato ingresado OK en el input -> %r", escrito)
    else:
        log.warning("[8/8] SIAC: el input quedó con %r en vez de %r", escrito, valor)

    # La lupa: botón con ícono de búsqueda al lado de la caja
    lupa = None
    for xp in ("//button[contains(@class,'search')]", "//*[contains(@class,'glyphicon-search')]/..",
               "//*[contains(@class,'fa-search')]/..", "//button[@type='submit']"):
        els = [e for e in driver.find_elements(By.XPATH, xp) if e.is_displayed()]
        if els:
            lupa = els[0]
            break

    if lupa is not None:
        log.info("[8/8] SIAC: botón de búsqueda encontrado; presionándolo...")
        try:
            lupa.click()
            log.info("[8/8] SIAC: BOTÓN PRESIONADO (clic normal).")
        except Exception as e:
            log.info("[8/8] SIAC: clic normal falló (%s); presionando por JavaScript...", type(e).__name__)
            driver.execute_script("arguments[0].click();", lupa)
            log.info("[8/8] SIAC: BOTÓN PRESIONADO (vía JavaScript).")
    else:
        log.info("[8/8] SIAC: no ubiqué el botón de la lupa; enviando ENTER en el input...")
        caja.send_keys(Keys.ENTER)
        log.info("[8/8] SIAC: ENTER enviado.")

    time.sleep(3)
    log.info("[8/8] SIAC: búsqueda enviada. URL actual: %s", driver.current_url)
    return True


def leer_valores(ruta=ARCHIVO_VALORES, columna=COLUMNA_VALORES, hoja=HOJA_VALORES):
    """Lee el Excel y devuelve los valores a buscar, sin vacíos ni duplicados."""
    from openpyxl import load_workbook

    archivo = Path(ruta)
    if not archivo.exists():                       # por si el nombre vino sin extensión
        candidatos = sorted(Path(".").glob(archivo.stem + ".xls*"))
        if not candidatos:
            raise RuntimeError(f"No encontré '{ruta}' en {Path('.').resolve()}")
        archivo = candidatos[0]

    wb = load_workbook(archivo, read_only=True, data_only=True)
    ws = wb[hoja] if isinstance(hoja, str) else wb.worksheets[hoja]
    filas = ws.iter_rows(values_only=True)
    cabecera = ["" if c is None else str(c).strip() for c in next(filas, ())]

    if isinstance(columna, int):
        idx = columna
    elif columna:
        idx = cabecera.index(columna)
    else:
        idx = next((i for i, h in enumerate(cabecera) if re.search(r"ruc|documento|dni", h, re.I)), None)
        if idx is None:
            idx = 0
            log.warning("[8/8] No identifiqué la columna; uso la primera. Encabezados: %s",
                        ", ".join(cabecera) or "(ninguno)")

    def limpiar(v):
        v = str(v).strip()
        return v[:-2] if v.endswith(".0") else v    # openpyxl devuelve los números como 20123456789.0

    valores, vistos = [], []
    primera = cabecera[idx] if idx < len(cabecera) else ""
    if primera.replace(" ", "").isdigit() and len(primera) >= 8:
        valores.append(limpiar(primera))            # el archivo no tenía encabezados

    for fila in filas:
        if idx >= len(fila) or fila[idx] is None:
            continue
        v = limpiar(fila[idx])
        if v and v not in vistos:
            vistos.append(v)
            valores.append(v)
    wb.close()

    log.info("[8/8] %s: %d valores en la columna %r", archivo.name, len(valores),
             cabecera[idx] if idx < len(cabecera) else idx)
    return valores


def _valor_celda(celda):
    """El número puede venir dentro de un <input> o como texto plano."""
    inputs = celda.find_elements(By.TAG_NAME, "input")
    if inputs:
        return (inputs[0].get_attribute("value") or "").strip()
    return celda.text.strip()


def _a_entero(texto):
    digitos = re.sub(r"[^0-9]", "", texto or "")
    return int(digitos) if digitos else None


def esperar_popup(driver, handles_antes, timeout=TIMEOUT_POPUP):
    """Devuelve el handle del popup de resultados, o None si no apareció."""
    fin = time.time() + timeout
    nuevo = None
    while time.time() < fin:
        for h in driver.window_handles:
            if h in handles_antes:
                continue
            try:
                driver.switch_to.window(h)
                nuevo = h
                if REGEX_POPUP.search(driver.current_url) or REGEX_POPUP.search(driver.title or ""):
                    log.info("[8/8] Popup detectado: %s", driver.current_url)
                    return h
            except Exception:
                continue          # todavía abriendo o ya se cerró
        time.sleep(1)

    if nuevo and nuevo in driver.window_handles:
        log.info("[8/8] Ventana nueva sin el patrón esperado; la uso igual.")
        return nuevo
    return None


def _leer_tablas(driver):
    """Una pasada sobre las tablas del popup: [(clave_cuenta, activas, no_activas), ...]."""
    filas_datos = []
    for i, tabla in enumerate(driver.find_elements(By.TAG_NAME, "table")):
        try:
            cabecera = [_normaliza(" ".join(th.text.split()))
                        for th in tabla.find_elements(By.TAG_NAME, "th")]
        except Exception:
            continue

        i_act = next((j for j, h in enumerate(cabecera) if "total activas" in h), None)
        i_no = next((j for j, h in enumerate(cabecera) if "total no activas" in h), None)
        if i_act is None or i_no is None:
            continue                       # no es una tabla de servicios
        i_cta = next((j for j, h in enumerate(cabecera) if "cuenta" in h), None)

        for fila in tabla.find_elements(By.XPATH, ".//tbody/tr"):
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) <= max(i_act, i_no):
                continue
            act = _a_entero(_valor_celda(celdas[i_act]))
            no = _a_entero(_valor_celda(celdas[i_no]))
            if act is None and no is None:
                continue
            cuenta = _valor_celda(celdas[i_cta]) if i_cta is not None else str(len(filas_datos))
            filas_datos.append((f"{i}|{cuenta}", act or 0, no or 0))

    return filas_datos


def extraer_totales(driver, timeout=TIMEOUT_DATOS):
    """Espera a que el popup termine de cargar sus tablas y recién ahí las lee.

    Relee cada pocos segundos: da por buena la lectura cuando dos pasadas seguidas
    devuelven la misma cantidad de filas, así no lee una tabla a medio pintar.
    """
    _esperar_carga(driver)
    log.info("[8/8] Esperando los datos del popup (hasta %s s)...", timeout)

    fin = time.time() + timeout
    filas, previas = [], -1
    while time.time() < fin:
        filas = _leer_tablas(driver)
        if filas and len(filas) == previas:
            log.info("[8/8] Datos completos: %d filas de servicios.", len(filas))
            return filas
        if filas:
            log.info("[8/8] Cargando... %d filas hasta ahora.", len(filas))
        previas = len(filas)
        time.sleep(ESPERA_DATOS)

    if filas:
        log.warning("[8/8] Se agotó el tiempo; me quedo con las %d filas leídas.", len(filas))
    return filas


def elementos_activas(driver, solo_positivos=True):
    """Los campos 'Total Activas' del popup, para hacerles clic: [(elemento, valor), ...]."""
    objetivos = []
    for tabla in driver.find_elements(By.TAG_NAME, "table"):
        try:
            cabecera = [_normaliza(" ".join(th.text.split()))
                        for th in tabla.find_elements(By.TAG_NAME, "th")]
        except Exception:
            continue
        i_act = next((j for j, h in enumerate(cabecera)
                      if "total activas" in h and "no activas" not in h), None)
        if i_act is None:
            continue
        for fila in tabla.find_elements(By.XPATH, ".//tbody/tr"):
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) <= i_act:
                continue
            celda = celdas[i_act]
            valor = _a_entero(_valor_celda(celda)) or 0
            if solo_positivos and valor <= 0:
                continue                      # un 0 no tiene detalle que mostrar
            campos = celda.find_elements(By.TAG_NAME, "input")
            objetivos.append((campos[0] if campos else celda, valor))
    return objetivos


def inspeccionar_popup(driver, max_filas=3):
    """Registra en el log qué trae el popup: título, columnas y primeras filas."""
    log.info("[8/8] Detalle -> %r | %s", driver.title, driver.current_url)
    for i, tabla in enumerate(driver.find_elements(By.TAG_NAME, "table")):
        try:
            cabecera = [" ".join(th.text.split()) for th in tabla.find_elements(By.TAG_NAME, "th")]
            filas = tabla.find_elements(By.XPATH, ".//tbody/tr")
        except Exception:
            continue
        if not cabecera and not filas:
            continue
        log.info("[8/8]   tabla %d (%d filas): %s", i, len(filas), " | ".join(cabecera) or "(sin encabezados)")
        for fila in filas[:max_filas]:
            celdas = [_valor_celda(c) for c in fila.find_elements(By.TAG_NAME, "td")]
            if any(celdas):
                log.info("[8/8]     %s", " | ".join(celdas))


def abrir_detalle_activas(driver, handle_popup):
    """Clic en el número de Total Activas; devuelve el handle del segundo popup."""
    objetivos = elementos_activas(driver)
    if not objetivos:
        log.info("[8/8] Ninguna fila con activas > 0; no hay detalle que abrir.")
        return None

    elemento, valor = objetivos[0]
    handles_antes = driver.window_handles
    log.info("[8/8] Clic en Total Activas = %d (de %d filas con activas)...", valor, len(objetivos))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
    time.sleep(0.5)
    try:
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)

    detalle = esperar_popup(driver, handles_antes)
    if detalle is None:
        log.warning("[8/8] El clic no abrió un segundo popup.")
        try:
            driver.switch_to.window(handle_popup)
        except Exception:
            pass
        return None

    driver.switch_to.window(detalle)
    _esperar_carga(driver)
    time.sleep(ESPERA_DATOS)
    inspeccionar_popup(driver)
    return detalle


def sumar_totales(filas):
    """Suma activas y no activas; agrupa por cuenta para no contar dos veces el mismo total."""
    if AGRUPAR_POR_CUENTA:
        por_cuenta = {}
        for clave, act, no in filas:
            por_cuenta[clave] = (act, no)
        pares = list(por_cuenta.values())
    else:
        pares = [(act, no) for _, act, no in filas]
    return sum(a for a, _ in pares), sum(n for _, n in pares)


def cerrar_popup(driver, handle_popup, handle_siac):
    """Clic en Cerrar; si no aparece el botón, cierra la ventana directamente."""
    try:
        xp = (_xpath_texto(("button", "a", "span"), "cerrar") +
              " | //input[contains(translate(@value,'CERRAR','cerrar'),'cerrar')]")
        botones = [e for e in driver.find_elements(By.XPATH, xp) if e.is_displayed()]
        if botones:
            log.info("[8/8] Cerrando el popup con el botón Cerrar.")
            try:
                botones[0].click()
            except Exception:
                driver.execute_script("arguments[0].click();", botones[0])
            time.sleep(1.5)
    except Exception as e:
        log.debug("Botón Cerrar no utilizable: %s", e)

    if handle_popup in driver.window_handles:      # el botón no la cerró
        try:
            driver.switch_to.window(handle_popup)
            driver.close()
            log.info("[8/8] Popup cerrado por el navegador.")
        except Exception:
            pass

    try:
        driver.switch_to.window(handle_siac)
    except Exception:
        vivas = _ventanas_vivas(driver)
        if vivas:
            driver.switch_to.window(vivas[-1][0])


def guardar_fila_csv(ruc, activas, no_activas, ruta=ARCHIVO_CSV):
    archivo = Path(ruta)
    nuevo = not archivo.exists()
    with open(archivo, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if nuevo:
            w.writerow(["ruc", "nro total activas", "nro total no activas"])
        w.writerow([ruc, activas, no_activas])


def buscar_lote(driver, valores=None):
    """Busca en SIAC cada valor del Excel, uno tras otro."""
    valores = leer_valores() if valores is None else valores
    if not valores:
        log.warning("[8/8] El Excel no tiene valores que buscar.")
        return

    url_siac = driver.current_url
    handle_siac = driver.current_window_handle
    total = len(valores)
    log.info("[8/8] Lote de %d búsquedas. Resultados -> %s", total, Path(ARCHIVO_CSV).resolve())

    for i, valor in enumerate(valores, 1):
        log.info("[8/8] (%d/%d) --- %s ---", i, total, valor)
        try:
            handles_antes = driver.window_handles
            buscar_en_siac(driver, valor)

            popup = esperar_popup(driver, handles_antes)
            if popup is None:
                log.warning("[8/8] (%d/%d) %s: no apareció el popup en %s s.", i, total, valor, TIMEOUT_POPUP)
                guardar_fila_csv(valor, "", "")
            else:
                driver.switch_to.window(popup)
                filas = extraer_totales(driver)
                if filas:
                    activas, no_activas = sumar_totales(filas)
                    log.info("[8/8] (%d/%d) %s -> activas=%d, no activas=%d (%d filas)",
                             i, total, valor, activas, no_activas, len(filas))
                    guardar_fila_csv(valor, activas, no_activas)
                    if CLIC_EN_TOTAL_ACTIVAS:
                        detalle = abrir_detalle_activas(driver, popup)
                        if detalle:
                            cerrar_popup(driver, detalle, popup)   # cierro el detalle, vuelvo al 1er popup
                else:
                    log.info("[8/8] (%d/%d) %s: el popup vino vacío.", i, total, valor)
                    guardar_fila_csv(valor, "", "")
                cerrar_popup(driver, popup, handle_siac)

        except Exception as e:
            log.warning("[8/8] (%d/%d) %s falló: %s. Recargo SIAC y sigo.", i, total, valor, e)
            guardar_fila_csv(valor, "ERROR", "ERROR")
            try:
                for h in driver.window_handles:      # cierro popups que hayan quedado sueltos
                    if h != handle_siac:
                        driver.switch_to.window(h)
                        driver.close()
                driver.switch_to.window(handle_siac)
                driver.get(url_siac)
                _esperar_carga(driver)
                seleccionar_tipo_busqueda(driver)
            except Exception as e2:
                log.error("[8/8] No pude recuperar la sesión: %s. Corto el lote.", e2)
                break
        time.sleep(PAUSA_ENTRE_BUSQUEDAS)

    log.info("[8/8] Lote terminado (%d valores). CSV: %s", total, Path(ARCHIVO_CSV).resolve())


def _ventanas_vivas(driver):
    """(handle, título, url) de las ventanas que todavía responden."""
    vivas = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            vivas.append((h, driver.title or "", driver.current_url or ""))
        except Exception:
            continue          # se cerró mientras la consultaba
    return vivas


def ir_a_ventana_siac(driver, pestanas_antes, timeout=TIMEOUT_VENTANA_SIAC):
    """Se cambia a la ventana de SIAC, tolerando popups que se abren y se cierran."""
    fin = time.time() + timeout
    nueva = None
    while time.time() < fin:
        vivas = _ventanas_vivas(driver)
        for h, titulo, url in vivas:
            if REGEX_VENTANA_SIAC.search(url) or REGEX_VENTANA_SIAC.search(titulo):
                driver.switch_to.window(h)
                log.info("[7/8] Ventana de SIAC: %r -> %s", titulo, url)
                return True
        nuevas = [v for v in vivas if v[0] not in pestanas_antes]
        if nuevas:
            nueva = nuevas[-1]     # existe pero aún no resuelve su URL
        time.sleep(1)

    vivas = _ventanas_vivas(driver)
    if not vivas:
        raise RuntimeError("Se cerraron todas las ventanas del navegador.")

    if nueva and any(v[0] == nueva[0] for v in vivas):
        driver.switch_to.window(nueva[0])
        log.warning("[7/8] Uso la ventana nueva %r; no confirmé que sea SIAC.", nueva[1])
        return False

    log.warning("[7/8] No apareció la ventana de SIAC en %s s. Ventanas vivas:", timeout)
    for _, titulo, url in vivas:
        log.warning("[7/8]   %r -> %s", titulo, url)
    driver.switch_to.window(vivas[-1][0])
    return False


def _esperar_carga(driver, timeout=TIMEOUT_WEB):
    """readyState == complete. False si la ventana murió (no lanza excepción)."""
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            if driver.execute_script("return document.readyState") == "complete":
                return True
        except Exception:
            return False
        time.sleep(0.5)
    return False


def clic_opcion_portal(driver, texto=TEXTO_OPCION):
    """Busca la opción (enlace, botón o texto) dentro del portal y le hace clic."""
    # Normaliza a minúsculas y sin tildes para comparar (el portal escribe "Unico" y "Único")
    minus = "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ", "abcdefghijklmnopqrstuvwxyzaeiouun"
    objetivo = texto.lower().translate(str.maketrans("áéíóúüñ", "aeiouun"))
    tr = f"translate(normalize-space(.),'{minus[0]}','{minus[1]}')"
    tr_attr = lambda a: f"translate(normalize-space(@{a}),'{minus[0]}','{minus[1]}')"
    xpath = (
        f"//a[contains({tr},'{objetivo}')]"
        f" | //button[contains({tr},'{objetivo}')]"
        f" | //span[contains({tr},'{objetivo}')]"
        f" | //td[contains({tr},'{objetivo}')]"
        f" | //li[contains({tr},'{objetivo}')]"
        f" | //*[contains({tr_attr('title')},'{objetivo}') or contains({tr_attr('alt')},'{objetivo}')]"
    )
    el = _buscar_en_frames(driver, xpath)

    if el is None:
        enlaces = sorted({a.text.strip() for a in driver.find_elements(By.TAG_NAME, "a")
                          if a.text.strip() and len(a.text.strip()) < 60})[:60]
        raise RuntimeError(
            f"No encontré la opción '{texto}' en el portal. "
            f"Enlaces visibles: {', '.join(enlaces) or 'ninguno (¿el portal cargó?)'}"
        )

    log.info("[7/8] Portal: clic en %r", el.text.strip() or texto)
    pestanas_antes = driver.window_handles
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)  # por si algo la tapa

    ir_a_ventana_siac(driver, pestanas_antes)

    # SIAC pide autenticación otra vez (es otra aplicación tras el portal)
    time.sleep(2)
    esperar_credenciales_usuario(driver)

    if not _esperar_carga(driver):
        # El popup intermedio se cerró al redirigir: busco la ventana definitiva
        log.warning("[7/8] La ventana se cerró mientras cargaba; busco la de SIAC otra vez.")
        ir_a_ventana_siac(driver, pestanas_antes)
        esperar_credenciales_usuario(driver)
        _esperar_carga(driver)
    log.info("[7/8] SIAC listo: %s", driver.current_url)

    time.sleep(2)                       # que SIAC termine de pintar la barra de búsqueda
    seleccionar_tipo_busqueda(driver)
    if VALOR_BUSQUEDA:
        buscar_en_siac(driver, VALOR_BUSQUEDA)     # override: un solo valor
    else:
        buscar_lote(driver)                        # itera el Excel


def _credenciales_portal():
    if REUSAR_CREDENCIALES_VPN:
        pwd = keyring.get_password(KEYRING_SERVICIO, USUARIO_VPN)   # la misma de la VPN
        origen = "--guardar-password"
    else:
        pwd = keyring.get_password(KEYRING_SERVICIO_PORTAL, USUARIO_PORTAL)
        origen = "--guardar-password-portal"
    if not pwd:
        raise RuntimeError(f"No hay contraseña guardada. Ejecuta: python toques_y_selenium.py {origen}")
    return USUARIO_PORTAL, pwd


def auth_basica(driver):
    """Manda las credenciales en la cabecera Authorization para evitar el popup del 401."""
    usuario, pwd = _credenciales_portal()
    token = base64.b64encode(f"{usuario}:{pwd}".encode()).decode()
    del pwd
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"Authorization": f"Basic {token}"}})
    log.info("Cabecera de autenticación configurada para el portal.")


def _ventana_navegador():
    for w in _escritorio().windows():
        try:
            marca = "Edge" if NAVEGADOR == "edge" else "Chrome"
            if w.class_name() == "Chrome_WidgetWin_1" and w.is_visible() and marca in (w.window_text() or ""):
                return w
        except Exception:
            continue
    return None


def _hay_dialogo_auth(driver):
    """El popup bloquea la carga: la página queda sin contenido."""
    try:
        return len(driver.find_elements(By.CSS_SELECTOR, "body *")) == 0
    except Exception:
        return True


def llenar_credenciales_lento(driver):
    """Escribe usuario y contraseña en el popup, tecla por tecla y con pausas largas."""
    ventana = _ventana_navegador()
    if ventana is None:
        log.warning("[6/8] No encontré la ventana del navegador.")
        return False

    usuario, pwd = _credenciales_portal()
    log.info("[6/8] Portal: tecleando credenciales de %s (%.2fs por tecla)...", usuario, PAUSA_TECLA)

    ventana.set_focus()
    time.sleep(ESPERA_FOCO)          # el diálogo tarda en tomar el foco

    for ch in usuario:               # tecla por tecla, no de golpe
        send_keys(_escapar_teclas(ch))
        time.sleep(PAUSA_TECLA)

    time.sleep(1)
    send_keys("{TAB}")
    time.sleep(1)

    for ch in pwd:
        send_keys(_escapar_teclas(ch))
        time.sleep(PAUSA_TECLA)
    del pwd

    time.sleep(1)
    send_keys("{ENTER}")
    log.info("[6/8] Portal: credenciales enviadas, esperando respuesta...")

    for _ in range(15):              # hasta 15s a que el portal responda
        if not _hay_dialogo_auth(driver):
            log.info("[6/8] Portal: credenciales ACEPTADAS.")
            return True
        time.sleep(1)
    return False


def esperar_credenciales_usuario(driver):
    """Si sale el popup de autenticación, lo llena automáticamente."""
    if not _hay_dialogo_auth(driver):
        log.info("[6/8] Portal: cargó sin pedir credenciales.")
        return True

    if llenar_credenciales_lento(driver):
        return True

    log.warning("[6/8] El login automático no funcionó; continúo igual.")
    return False


def verificar_password():
    """Diagnostica el esquema de autenticación del portal y prueba la contraseña guardada."""
    guardada = keyring.get_password(KEYRING_SERVICIO, USUARIO_VPN)
    if not guardada:
        print(f"No hay contraseña guardada para {USUARIO_VPN}.")
        print("Ejecuta: python toques_y_selenium.py --guardar-password")
        return

    print(f"Guardada para {USUARIO_VPN}: {guardada[0]}{'*' * (len(guardada) - 2)}{guardada[-1]}"
          f"  ({len(guardada)} caracteres)")

    try:
        import requests
        from requests.auth import HTTPBasicAuth
    except ImportError:
        print("Instala requests: pip install requests")
        return

    # 1) ¿Qué esquema pide el servidor?
    try:
        r = requests.get(URL_SELENIUM, timeout=20)
    except Exception as e:
        print(f"No se pudo conectar ({e}). ¿Está la VPN arriba?")
        return

    esquemas = r.headers.get("WWW-Authenticate", "")
    print(f"El portal responde {r.status_code}; esquemas que pide: {esquemas or '(ninguno)'}")

    usa_ntlm = "ntlm" in esquemas.lower() or "negotiate" in esquemas.lower()
    if usa_ntlm:
        print("=> Usa NTLM/Negotiate, no Basic. Por eso la cabecera Basic no funciona.")

    # 2) Probar las credenciales con el esquema correcto
    intentos = []
    if usa_ntlm:
        try:
            from requests_ntlm import HttpNtlmAuth
        except ImportError:
            print("Para probar NTLM: pip install requests-ntlm")
            return
        for usr in (USUARIO_PORTAL, f"TIM\\{USUARIO_VPN}", f"{USUARIO_VPN}@tim.com.pe"):
            intentos.append((usr, HttpNtlmAuth(usr, guardada)))
    else:
        intentos.append((USUARIO_PORTAL, HTTPBasicAuth(USUARIO_PORTAL, guardada)))

    for usr, auth in intentos:
        try:
            r = requests.get(URL_SELENIUM, auth=auth, timeout=30)
        except Exception as e:
            print(f"  {usr}: error ({e})")
            continue
        if r.status_code < 400:
            print(f"  {usr}: ACEPTADO ({r.status_code}). Usa este valor en USUARIO_PORTAL.")
            return
        print(f"  {usr}: rechazado ({r.status_code})")

    print("Ninguna variante funcionó. Puede ser un dominio distinto al que asumí ('TIM').")


def guardar_password_portal():
    import getpass
    pwd = getpass.getpass(f"Contraseña del portal para {USUARIO_PORTAL} (no se mostrará): ")
    if not pwd:
        print("Contraseña vacía; no se guardó nada.")
        return
    keyring.set_password(KEYRING_SERVICIO_PORTAL, USUARIO_PORTAL, pwd)
    print("Guardada en el Administrador de credenciales de Windows.")


def flujo_portal(url=URL_SELENIUM, opcion=TEXTO_OPCION):
    """Solo la parte de Selenium: abre el portal, clica la opción y devuelve el driver.

    Uso independiente (con la VPN ya conectada):
        from toques_y_selenium import flujo_portal
        driver = flujo_portal()
    """
    driver = abrir_selenium(url, navegar=False)
    driver.get(url)                        # con pageLoadStrategy 'none' vuelve enseguida
    esperar_credenciales_usuario(driver)
    _esperar_carga(driver)
    log.info("[6/8] Portal listo: %s", driver.current_url)
    clic_opcion_portal(driver, opcion)
    return driver


# ---------------------------- main ----------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--listar", action="store_true", help="Solo listar las apps de iconos ocultos (sin Selenium)")
    ap.add_argument("--solo-selenium", action="store_true", help="Saltar la bandeja y abrir solo Selenium")
    ap.add_argument("--capturar-ivanti", action="store_true",
                    help="Mostrar coordenadas del mouse relativas a la ventana de Ivanti")
    ap.add_argument("--guardar-password", action="store_true",
                    help="Guardar la contraseña de la VPN en el Administrador de credenciales de Windows")
    ap.add_argument("--guardar-password-portal", action="store_true",
                    help="Guardar la contraseña del portal de distribuidores")
    ap.add_argument("--verificar-password", action="store_true",
                    help="Ver qué contraseña hay guardada y probarla contra el portal")
    args = ap.parse_args()

    try:
        if args.verificar_password:
            verificar_password()
            sys.exit(0)

        if args.guardar_password:
            guardar_password()
            sys.exit(0)

        if args.guardar_password_portal:
            guardar_password_portal()
            sys.exit(0)

        if args.capturar_ivanti:
            capturar_coords_ivanti()
            sys.exit(0)

        if args.listar:
            listar_iconos_ocultos()
            sys.exit(0)

        if not args.solo_selenium:
            clic_icono_oculto()
            clic_conectar_ivanti()
            log.info("Esperando %s s antes de abrir Selenium...", ESPERA_ANTES_SELENIUM)
            time.sleep(ESPERA_ANTES_SELENIUM)

        driver = flujo_portal()

        # ---- Aquí sigue tu lógica dentro de SIAC ----

        if not MANTENER_ABIERTO:
            input("Presiona Enter para cerrar el navegador...")
            driver.quit()

    except Exception as e:
        log.exception("Error: %s", e)
        sys.exit(1)
