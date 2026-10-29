"""
funciones_isams.py  -  Biblioteca de funciones compartidas (iSAMS, colegio nuevo)

Aquí viven las utilidades reutilizables (shadow DOM, scroll del árbol
virtualizado, alertas nativas, driver y login). Los scripts principales
(leer_materias.py, etc.) solo importan de aquí y de datos.py.
"""

import re
import time
import unicodedata

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import UnexpectedAlertPresentException

from datos import CHROME_PATH, DRIVER_PATH, LOGIN_URL, COMPARTIR_CALCULADAS, CLASIFICACION_POR_GRADO, ALIAS_CARPETAS


# --- Navegador y sesión ---

def crear_driver():
    """Abre Chrome (Selenium Manager resuelve el driver si DRIVER_PATH es None)."""
    options = webdriver.ChromeOptions()
    if CHROME_PATH:
        options.binary_location = CHROME_PATH
    if DRIVER_PATH:
        driver = webdriver.Chrome(service=Service(DRIVER_PATH), options=options)
    else:
        driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def login_manual(driver):
    """Abre la URL de login y espera a que el usuario inicie sesión a mano."""
    driver.get(LOGIN_URL)
    input("\n>>> Inicia sesión manualmente en la ventana del navegador.\n"
          ">>> Cuando ya estés dentro de la plataforma, vuelve aquí y presiona ENTER...\n")


# --- Búsqueda en DOM / shadow DOM ---


_DEEP_QUERY_JS = """
function recolectar(root, selector, resultados) {
    resultados.push(...root.querySelectorAll(selector));
    const todos = root.querySelectorAll('*');
    for (const nodo of todos) {
        if (nodo.shadowRoot) {
            recolectar(nodo.shadowRoot, selector, resultados);
        }
    }
}
const resultados = [];
recolectar(document, arguments[0], resultados);
// Esta plataforma reutiliza ids genericos (form, input, position, wizard...)
// en muchos componentes, y los dialogos de Kendo suelen quedar ocultos en el
// DOM al cerrarse en vez de eliminarse. Por eso, entre varias coincidencias,
// preferimos la que este realmente visible en pantalla (la actual), no la
// primera que aparezca en el orden del documento (que podria ser un dialogo
// viejo ya cerrado).
for (const el of resultados) {
    if (el.offsetParent !== null || el.getClientRects().length > 0) {
        return el;
    }
}
return resultados.length > 0 ? resultados[0] : null;
"""


_ENCONTRAR_CONTENEDOR_SCROLLEABLE_JS = """
function encontrarContenedorScrolleable(el) {
    let nodo = el;
    while (nodo && nodo !== document.body) {
        if (nodo.scrollHeight > nodo.clientHeight + 5) {
            return nodo;
        }
        nodo = nodo.parentElement || (nodo.getRootNode() && nodo.getRootNode().host);
    }
    return document.scrollingElement || document.documentElement;
}
const ancla = document.querySelector('div.row');
if (!ancla) return null;
return encontrarContenedorScrolleable(ancla);
"""


def elemento_profundo(driver, selector, tiempo_espera=20):
    """Busca un elemento en cualquier parte del documento, atravesando
    cualquier cantidad de shadow roots anidados."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        el = driver.execute_script(_DEEP_QUERY_JS, selector)
        if el is not None:
            return el
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró el elemento (búsqueda profunda): {selector}")


def elemento_profundo_por_texto(driver, selector_contenedor, texto_exacto, tiempo_espera=20):
    """Como elemento_profundo, pero filtra por el texto visible exacto del
    elemento (útil cuando varios elementos comparten la misma clase), y
    entre los que coincidan prefiere el que esté realmente visible en
    pantalla (por los mismos diálogos viejos ocultos que motivan
    elemento_profundo)."""
    script = """
        function recolectar(root, selector, resultados) {
            resultados.push(...root.querySelectorAll(selector));
            const todos = root.querySelectorAll('*');
            for (const nodo of todos) {
                if (nodo.shadowRoot) {
                    recolectar(nodo.shadowRoot, selector, resultados);
                }
            }
        }
        const candidatos = [];
        recolectar(document, arguments[0], candidatos);
        const coincidencias = candidatos.filter(
            (c) => (c.textContent || '').trim() === arguments[1]
        );
        for (const el of coincidencias) {
            if (el.offsetParent !== null || el.getClientRects().length > 0) {
                return el;
            }
        }
        return coincidencias.length > 0 ? coincidencias[0] : null;
    """
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        el = driver.execute_script(script, selector_contenedor, texto_exacto)
        if el is not None:
            return el
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró '{texto_exacto}' con selector: {selector_contenedor}")


def dentro_de_shadow(driver, host_element, selector_interno, tiempo_espera=20):
    """Dado un elemento host ya localizado, busca un elemento dentro de su
    propio shadow root (un solo nivel, no recursivo)."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        el = driver.execute_script(
            "return arguments[0].shadowRoot ? arguments[0].shadowRoot.querySelector(arguments[1]) : null;",
            host_element, selector_interno,
        )
        if el is not None:
            return el
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró dentro del shadow root: {selector_interno}")


def dentro_de_shadow_todos(driver, host_element, selector_interno):
    """Como dentro_de_shadow, pero devuelve TODOS los elementos que matchean
    dentro del shadow root del host (un solo nivel), no solo el primero."""
    return driver.execute_script(
        "return arguments[0].shadowRoot ? Array.from(arguments[0].shadowRoot.querySelectorAll(arguments[1])) : [];",
        host_element, selector_interno,
    )


def clic_seguro(driver, elemento):
    """Intenta un clic normal; si falla, dispara el clic vía JavaScript."""
    try:
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)


def clic_unico_para_alerta(driver, elemento):
    """Da UN solo clic, sin reintentar con un segundo clic de respaldo (a
    diferencia de clic_seguro). Se usa para elementos que disparan un
    alert() nativo síncrono: si reintentáramos con un clic de JS tras la
    excepción del primero, dispararíamos el alert DOS veces, dejando uno
    huérfano bloqueando toda la sesión de Selenium."""
    try:
        elemento.click()
    except UnexpectedAlertPresentException:
        pass  # esperado: el alert ya se abrió durante este mismo clic


def cerrar_alerta_pendiente(driver):
    """Si por algún motivo quedó un alert nativo abierto (de un paso
    anterior que falló a mitad de camino), lo cierra antes de seguir, para
    que no bloquee todos los comandos de Selenium siguientes."""
    try:
        alerta = driver.switch_to.alert
        alerta.accept()
        return True
    except Exception:
        return False


def xpath_literal(texto):
    """Literal seguro para XPath, incluso si el texto tiene comillas."""
    if "'" not in texto:
        return f"'{texto}'"
    if '"' not in texto:
        return f'"{texto}"'
    partes = texto.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in partes) + ")"


def normalizar(texto):
    """Mayúsculas, sin tildes, espacios colapsados, para comparar textos con
    formato inconsistente (MAYÚSCULAS/Capitalizado, con/sin tildes, etc.)."""
    texto = (texto or "").strip().upper()
    texto = "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))
    texto = " ".join(texto.split())
    return texto


def buscar_ligero_desde(driver, elemento_raiz, selector, tiempo_espera=20):
    """Busca dentro del DOM ligero (sin cruzar shadow) de un elemento ya
    localizado, usando querySelector normal escopado a ese elemento. Útil
    para evitar ambigüedad con ids genéricos reutilizados en otras partes
    de la página, una vez que ya tenemos el contenedor correcto (ej. el
    wizard actualmente abierto)."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        el = driver.execute_script("return arguments[0].querySelector(arguments[1]);", elemento_raiz, selector)
        if el is not None:
            return el
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró (DOM ligero) dentro del elemento: {selector}")


def buscar_elementos_con_scroll(driver, xpath, filtro=None, tiempo_espera=45):
    """Busca elementos por XPath en el árbol de estructura de columnas
    (markbooks-structure), que resultó ser VIRTUALIZADO: solo las filas
    visibles en pantalla existen en el DOM. Si no aparecen coincidencias,
    hace scroll incremental dentro del contenedor scrolleable más cercano
    (ubicado dinámicamente a partir de cualquier fila conocida) y reintenta,
    hasta encontrar algo o llegar al final del scroll (y de vuelta al
    inicio, por si la fila buscada quedó arriba). Devuelve una lista
    (puede estar vacía si no se encuentra nada)."""
    fin = time.time() + tiempo_espera
    ultimo_scroll_top = -1
    reinicio_probado = False

    while time.time() < fin:
        candidatos = driver.find_elements(By.XPATH, xpath)
        if filtro:
            candidatos = [c for c in candidatos if filtro(c)]
        if candidatos:
            return candidatos

        contenedor = driver.execute_script(_ENCONTRAR_CONTENEDOR_SCROLLEABLE_JS)
        if contenedor is None:
            time.sleep(0.3)
            continue

        nuevo_scroll_top = driver.execute_script(
            "arguments[0].scrollTop += arguments[0].clientHeight * 0.8; return arguments[0].scrollTop;",
            contenedor,
        )
        if nuevo_scroll_top == ultimo_scroll_top:
            if reinicio_probado:
                break  # ya dimos toda la vuelta (fin -> inicio) sin encontrar nada
            # Llegamos al final del scroll sin encontrar nada; probamos
            # también desde el inicio por si la fila buscada quedó arriba.
            driver.execute_script("arguments[0].scrollTop = 0;", contenedor)
            reinicio_probado = True
            ultimo_scroll_top = 0
            time.sleep(0.3)
            continue
        ultimo_scroll_top = nuevo_scroll_top
        time.sleep(0.3)

    return []


def nombres_carpeta(nombre_carpeta):
    """Conjunto de nombres normalizados que valen para esa carpeta: ella misma y sus alias (datos.ALIAS_CARPETAS)."""
    nombres = {normalizar(nombre_carpeta)}
    for clave, alias in ALIAS_CARPETAS.items():
        if normalizar(clave) == normalizar(nombre_carpeta):
            nombres.update(normalizar(a) for a in alias)
    return nombres


def buscar_id_carpeta(driver, nombre_carpeta, tiempo_espera=45):
    """Devuelve el id numérico de la fila-carpeta (tiene el atributo 'group')
    con ese título. Compara de forma insensible a mayúsculas/tildes, y
    scrollea el árbol (virtualizado) si hace falta hasta encontrarla."""
    objetivos = nombres_carpeta(nombre_carpeta)
    filas = buscar_elementos_con_scroll(
        driver,
        "//div[contains(@class,'row')][@group]",
        filtro=lambda f: normalizar(f.get_attribute("title") or "") in objetivos,
        tiempo_espera=tiempo_espera,
    )
    if not filas:
        raise TimeoutError(f"No se encontró la carpeta '{nombre_carpeta}' en el árbol de estructura.")
    try:
        return int(filas[0].get_attribute("id"))
    except (TypeError, ValueError):
        raise RuntimeError(f"La carpeta '{nombre_carpeta}' no tiene un id numérico válido.")


# --- Lectura del árbol de estructura de un gradebook (solo lectura) ---

_JS_LEER_FILAS = """
// Recolecta TODOS los textos de hoja (leaf) de una fila, atravesando shadow roots,
// en orden de aparición: Name | Short Name | Maximum | Weighting (%) ...
function textos(nodo, salida) {
    for (const hijo of nodo.childNodes) {
        if (hijo.nodeType === 3) {
            const t = (hijo.textContent || '').replace(/\\s+/g, ' ').trim();
            if (t) salida.push(t);
        } else if (hijo.nodeType === 1) {
            if (hijo.tagName === 'STYLE' || hijo.tagName === 'SCRIPT') continue;
            if (hijo.shadowRoot) textos(hijo.shadowRoot, salida);
            textos(hijo, salida);
        }
    }
}
function icono(el) {
    const ic = el.querySelector('isams-icon, i, svg, img');
    if (!ic) return '';
    return [ic.getAttribute('title'), ic.getAttribute('icon'), ic.getAttribute('name'),
            ic.getAttribute('class'), ic.getAttribute('src')]
        .filter(Boolean).join(' / ');
}
return Array.from(document.querySelectorAll('div.row')).map(function (el) {
    const t = [];
    textos(el, t);
    return {
        id: el.getAttribute('id') || '',
        titulo: el.getAttribute('title') || '',
        parent: el.getAttribute('parent') || '',
        carpeta: el.hasAttribute('group'),
        textos: t,
        icono: icono(el)
    };
});
"""

_JS_HTML_FILAS = """
return Array.from(document.querySelectorAll('div.row')).slice(0, arguments[0])
    .map(function (el) { return el.outerHTML; });
"""


def html_primeras_filas(driver, n=3):
    """outerHTML de las primeras filas (para diagnosticar el DOM si algo no se lee bien)."""
    return driver.execute_script(_JS_HTML_FILAS, n)


def analizar_fila(f):
    """Devuelve (nombre_corto, maximo, ponderacion, porcentaje) de una fila.
    Orden real de los textos en el DOM (verificado): el Name va primero y el
    Short Name es el ÚLTIMO texto de la fila; el Maximum (si existe) es el
    número que lo precede. '-' = vacío. 'porcentaje' es el número que acompaña
    al '%' dentro del nombre corto (ej. 'Calculo Insumo 70% T-I' -> 70)."""
    textos = list(f.get("textos") or [])
    nombre = (f.get("titulo") or "").strip()
    if nombre and nombre in textos:
        textos.remove(nombre)
    elif textos:
        textos.pop(0)

    corto = textos[-1].strip() if textos else ""
    if corto == "-":
        corto = ""
    maximo = ""
    if corto and len(textos) >= 2 and re.fullmatch(r"\d+(?:[.,]\d+)?", textos[-2].strip()):
        maximo = textos[-2].strip()

    m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", corto)
    porcentaje = float(m.group(1).replace(",", ".")) if m else None
    if porcentaje is not None and porcentaje == int(porcentaje):
        porcentaje = int(porcentaje)
    return corto, maximo, "", porcentaje


def filas_de_carpeta(filas, nombre_carpeta):
    """De {fila_id: fila} devuelve [(fila_id, fila_dict)] de las COLUMNAS dentro de la
    carpeta 'nombre_carpeta' (ej. 'TI'), en orden de aparición. Usa el atributo parent; si no
    hubiera, cae al orden del documento (todo lo que sigue a la carpeta hasta la siguiente
    carpeta). Devuelve None si la carpeta no existe."""
    objetivos = nombres_carpeta(nombre_carpeta)
    carpetas = [fid for fid, f in filas.items() if f["carpeta"] and normalizar(f["titulo"]) in objetivos]
    if not carpetas:
        return None
    carpeta_id = carpetas[0]

    hijas = [fid for fid, f in filas.items() if not f["carpeta"] and f["parent"] == carpeta_id]
    if not hijas:  # respaldo por orden
        hijas, dentro = [], False
        for fid, f in filas.items():
            if fid == carpeta_id:
                dentro = True
                continue
            if dentro and f["carpeta"]:
                break
            if dentro:
                hijas.append(fid)
    return [(fid, filas[fid]) for fid in hijas]


def columnas_de_carpeta(filas, nombre_carpeta):
    """[(fila_id, titulo, porcentaje)] de las columnas de la carpeta (None si no existe)."""
    hijas = filas_de_carpeta(filas, nombre_carpeta)
    if hijas is None:
        return None
    return [(fid, f["titulo"], analizar_fila(f)[3]) for fid, f in hijas]


def esperar_filas(driver, tiempo_espera=40):
    """Espera a que el árbol de estructura tenga al menos una div.row."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        try:
            if driver.execute_script("return document.querySelectorAll('div.row').length;") > 0:
                return True
        except UnexpectedAlertPresentException:
            cerrar_alerta_pendiente(driver)
        time.sleep(0.5)
    return False


def leer_estructura(driver, tiempo_espera=90):
    """Recorre el árbol (virtualizado) con scroll y devuelve {fila_id: dict}
    con id, titulo, parent y carpeta (True si es carpeta)."""
    filas = {}
    contenedor = driver.execute_script(_ENCONTRAR_CONTENEDOR_SCROLLEABLE_JS)
    if contenedor is not None:
        driver.execute_script("arguments[0].scrollTop = 0;", contenedor)
        time.sleep(0.4)

    ultimo_top = -1
    sin_cambio = 0
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        for f in driver.execute_script(_JS_LEER_FILAS):
            if f["id"]:
                filas[f["id"]] = f
        if contenedor is None:
            break
        top = driver.execute_script(
            "arguments[0].scrollTop += arguments[0].clientHeight * 0.8; return arguments[0].scrollTop;",
            contenedor,
        )
        if top == ultimo_top:
            sin_cambio += 1
            if sin_cambio >= 2:
                break
        else:
            sin_cambio = 0
        ultimo_top = top
        time.sleep(0.3)

    # Última pasada tras llegar al fondo.
    for f in driver.execute_script(_JS_LEER_FILAS):
        if f["id"]:
            filas[f["id"]] = f
    return filas


# --- Edición de columnas calculadas (wizard 'Edit Calculated Column') ---


def buscar_icono_en_fila(driver, row_element, clase_icono, tiempo_espera=20):
    """Dado un <div class='row'> del árbol de estructura, ubica un ícono
    específico (por su clase, ej. 'icon-16-column-edit' o
    'icon-16-branch-element') dentro del shadow DOM de algún isams-icon
    hijo, sin importar en qué posición esté entre los demás iconos de la
    fila, y devuelve el host (isams-icon) al que hay que darle clic."""
    script = """
        function buscarProfundoDesde(root, selector) {
            let el = root.querySelector(selector);
            if (el) return el;
            const todos = root.querySelectorAll('*');
            for (const nodo of todos) {
                if (nodo.shadowRoot) {
                    el = buscarProfundoDesde(nodo.shadowRoot, selector);
                    if (el) return el;
                }
            }
            return null;
        }
        const icono = buscarProfundoDesde(arguments[0], 'i.' + arguments[1]);
        if (!icono) return null;
        const raiz = icono.getRootNode();
        return raiz.host || icono;
    """
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        host = driver.execute_script(script, row_element, clase_icono)
        if host is not None:
            return host
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró el ícono '{clase_icono}' en la fila.")


def buscar_boton_editar_columna(driver, row_element, tiempo_espera=20):
    """Ubica el ícono de editar (lápiz) de una fila del árbol de estructura."""
    return buscar_icono_en_fila(driver, row_element, "icon-16-column-edit", tiempo_espera)


def clic_boton_wizard(driver, wizard_selector, texto_boton, tiempo_espera=20):
    """Dentro del shadow root del wizard, busca entre los isams-button de
    #wizardButtons el que tenga ese texto visible (insensible a mayúsculas)
    y le da clic."""
    wizard = elemento_profundo(driver, wizard_selector, tiempo_espera)
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        botones_host = dentro_de_shadow_todos(driver, wizard, "#wizardButtons isams-button")
        for host in botones_host:
            try:
                btn = driver.execute_script(
                    "return arguments[0].shadowRoot ? arguments[0].shadowRoot.querySelector('button.button') : null;",
                    host,
                )
                if btn is None:
                    continue
                texto = (btn.get_attribute("textContent") or "").strip()
            except Exception:
                continue
            if texto.lower() == texto_boton.lower():
                # Durante la transición entre pasos, el botón puede quedar
                # deshabilitado unos instantes mientras el contenido del
                # siguiente paso termina de cargar de forma asíncrona.
                try:
                    if not (btn.is_displayed() and btn.is_enabled()):
                        continue
                except Exception:
                    continue
                clic_seguro(driver, btn)
                return btn
        time.sleep(0.3)
    raise TimeoutError(f"No se encontró el botón del wizard con texto '{texto_boton}'")


def clic_next_wizard_verificado(driver, wizard_selector, tiempo_espera=20):
    """Como clic_boton_wizard con texto 'Next', pero además verifica que el
    indicador 'Step X of Y' del wizard realmente haya cambiado antes de
    continuar (por si el botón queda deshabilitado momentáneamente mientras
    el siguiente paso carga contenido de forma asíncrona, ej. una grilla de
    'Report Mappings')."""
    wizard = elemento_profundo(driver, wizard_selector, tiempo_espera)
    paging = dentro_de_shadow(driver, wizard, "#wizardPaging", tiempo_espera)
    texto_antes = (paging.get_attribute("textContent") or "").strip()

    fin = time.time() + tiempo_espera
    while time.time() < fin:
        try:
            clic_boton_wizard(driver, wizard_selector, "Next", tiempo_espera=5)
        except TimeoutError:
            pass  # el botón puede seguir deshabilitado; reintentamos abajo

        sub_fin = time.time() + 5
        while time.time() < sub_fin:
            wizard_actual = elemento_profundo(driver, wizard_selector, tiempo_espera)
            paging_actual = dentro_de_shadow(driver, wizard_actual, "#wizardPaging", tiempo_espera)
            texto_actual = (paging_actual.get_attribute("textContent") or "").strip()
            if texto_actual != texto_antes:
                return texto_actual
            time.sleep(0.3)
        # No cambió de paso todavía: volvemos a intentar el clic en Next

    raise TimeoutError(f"El wizard no avanzó de '{texto_antes}' tras reintentar 'Next' varias veces.")


def configurar_propiedades_calculo(driver, tiempo_espera=20):
    """Step 1 del wizard: Calculation Type = 'Formula Builder',
    Allow Sharing = COMPARTIR_CALCULADAS (datos.py), y avanza al Step 2."""
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)

    dropdown_calculo = dentro_de_shadow(driver, form, "isams-dropdown#calculation", tiempo_espera)
    select_calculo = dentro_de_shadow(driver, dropdown_calculo, "select#input", tiempo_espera)
    Select(select_calculo).select_by_visible_text("Formula Builder")

    dropdown_compartir = dentro_de_shadow(driver, form, "isams-dropdown#shareable", tiempo_espera)
    select_compartir = dentro_de_shadow(driver, dropdown_compartir, "select#input", tiempo_espera)
    Select(select_compartir).select_by_visible_text(COMPARTIR_CALCULADAS)

    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)


def abrir_pestana_columns(driver, tiempo_espera=20):
    """En el Step 2 (Formula Builder), clica la pestaña 'Columns' para que
    cargue la lista de columnas disponibles (con sus IDs). Escopada al
    wizard actual (isams-wizard#calculationWizard): tanto el id 'listitems'
    como el texto 'Columns' se repiten en otras partes de la página (el
    árbol principal detrás del modal), así que buscar en todo el documento
    puede agarrar el elemento equivocado (tapado visualmente, pero técnicamente
    'visible' para el navegador)."""
    wizard = elemento_profundo(driver, "isams-wizard#calculationWizard", tiempo_espera)
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        tab = driver.execute_script(
            """
            const spans = arguments[0].querySelectorAll('span.k-link');
            for (const s of spans) {
                if ((s.textContent || '').trim() === 'Columns') return s;
            }
            return null;
            """,
            wizard,
        )
        if tab is not None:
            clic_seguro(driver, tab)
            time.sleep(1)
            return
        time.sleep(0.3)
    raise TimeoutError("No se encontró la pestaña 'Columns' dentro del wizard actual.")


def recolectar_ids_columnas(driver, nombre_columna, tiempo_espera=60):
    """Dentro de la pestaña 'Columns' (markbooks-structure-display, también
    virtualizada), recolecta los IDs de TODAS las filas cuyo nombre (texto
    del div interno, ej. 'TOTAL Q1 PARCIAL 1') coincida exactamente,
    scrolleando desde el inicio hasta el final de la lista para no
    perderse ninguna (puede haber una por cada materia del gradebook).
    Cada fila es <div id="ID12345"><div>Nombre</div></div>. Escopado al
    wizard actual (isams-wizard#calculationWizard) por el mismo motivo que
    abrir_pestana_columns: el id 'listitems' se repite en el árbol de fondo.
    Devuelve una lista de ids SIN el prefijo 'ID' (ej. ['12345', '67890'])."""
    wizard = elemento_profundo(driver, "isams-wizard#calculationWizard", tiempo_espera)
    contenedor = driver.execute_script(
        "return arguments[0].querySelector("
        "\"markbooks-structure-display[id='columnSelection'] isams-pane[id='listitems']\""
        ");",
        wizard,
    )
    if contenedor is None:
        raise TimeoutError(
            "No se encontró la lista de columnas ('listitems' dentro de "
            "markbooks-structure-display#columnSelection) en el wizard actual."
        )

    # Esperar a que el panel realmente termine de renderizar contenido: justo
    # tras clicar la pestaña "Columns" puede tardar un instante en poblarse;
    # si consultamos demasiado pronto vemos el contenedor vacío.
    fin_espera_inicial = time.time() + 15
    while time.time() < fin_espera_inicial:
        hay_contenido = driver.execute_script(
            "return arguments[0].querySelectorAll(\"div[id^='ID'], div[id^='GID']\").length > 0;",
            contenedor,
        )
        if hay_contenido:
            break
        time.sleep(0.3)

    driver.execute_script("arguments[0].scrollTop = 0;", contenedor)
    time.sleep(0.3)

    ids_encontrados = {}
    textos_vistos_muestra = []
    ids_muestra_vistos = set()
    total_filas_vistas = 0
    objetivo_norm = normalizar(nombre_columna)
    ultimo_scroll_top = -1
    sin_cambio_seguidas = 0
    fin = time.time() + tiempo_espera

    while time.time() < fin:
        filas = driver.execute_script(
            "return Array.from(arguments[0].querySelectorAll(\"div[id^='ID']\"));",
            contenedor,
        )
        for fila in filas:
            id_attr = fila.get_attribute("id")
            if not id_attr:
                continue
            texto_completo = driver.execute_script("return arguments[0].textContent;", fila) or ""
            # El textContent de la fila trae el Nombre Y el ID (la columna
            # "ID" de la tabla también es texto dentro del mismo div), ej.
            # "TOTAL Q1 PARCIAL 1\n   \n   \nID217902". Nos quedamos solo con
            # la primera línea no vacía (el nombre).
            lineas_no_vacias = [l.strip() for l in texto_completo.split("\n") if l.strip()]
            texto = lineas_no_vacias[0] if lineas_no_vacias else ""
            total_filas_vistas += 1
            if id_attr not in ids_muestra_vistos and len(textos_vistos_muestra) < 15:
                textos_vistos_muestra.append((id_attr, texto))
                ids_muestra_vistos.add(id_attr)
            if id_attr in ids_encontrados:
                continue
            if normalizar(texto) == objetivo_norm:
                ids_encontrados[id_attr] = texto

        info_scroll = driver.execute_script(
            """
            const c = arguments[0];
            const puedeScrollear = c.scrollHeight > c.clientHeight + 5;
            if (puedeScrollear) {
                c.scrollTop += c.clientHeight * 0.8;
            }
            return {scrollTop: c.scrollTop, puedeScrollear: puedeScrollear};
            """,
            contenedor,
        )
        nuevo_scroll_top = info_scroll["scrollTop"]

        if not info_scroll["puedeScrollear"]:
            # Contenedor todavía sin tamaño real (recién abierta la
            # pestaña) o ya no hace falta scroll; esperamos un poco más en
            # vez de asumir que ya llegamos al final.
            time.sleep(0.3)
            continue

        if nuevo_scroll_top == ultimo_scroll_top:
            sin_cambio_seguidas += 1
            if sin_cambio_seguidas >= 3:
                break  # llegamos al final de la lista de verdad
        else:
            sin_cambio_seguidas = 0
        ultimo_scroll_top = nuevo_scroll_top
        time.sleep(0.2)

    if not ids_encontrados:
        print(f"   [diagnóstico] Se revisaron {total_filas_vistas} filas (con repeticiones por scroll) "
              f"buscando '{nombre_columna}'. Primeros textos vistos, con su id:")
        for id_attr, texto in textos_vistos_muestra:
            print(f"      {id_attr!r} -> {texto!r}")

    return [id_attr[2:] for id_attr in ids_encontrados.keys()]  # sin el prefijo "ID"


_JS_ESCRIBIR_FORMULA = """
// Busca (atravesando shadow roots) los isams-textarea#customFormula, descarta los que no están
// visibles (diálogos viejos que Kendo deja ocultos en el DOM) y escribe en el textarea del visible.
function recolectar(root, selector, resultados) {
    resultados.push(...root.querySelectorAll(selector));
    for (const nodo of root.querySelectorAll('*')) {
        if (nodo.shadowRoot) recolectar(nodo.shadowRoot, selector, resultados);
    }
}
const candidatos = [];
recolectar(document, 'isams-textarea#customFormula', candidatos);
for (const el of candidatos) {
    if (el.offsetParent === null && el.getClientRects().length === 0) continue;
    const ta = el.shadowRoot ? el.shadowRoot.querySelector('textarea#input') : null;
    if (!ta) continue;
    ta.value = arguments[0];
    ta.dispatchEvent(new Event('input', {bubbles: true}));
    ta.dispatchEvent(new Event('change', {bubbles: true}));
    return ta.value;
}
return null;
"""


def escribir_formula(driver, formula_texto, tiempo_espera=30):
    """Escribe (reemplazando el contenido) la fórmula en el textarea del Formula Builder
    del wizard ACTUAL. Reintenta hasta que aparezca el textarea VISIBLE: tras abrir un
    2.º/3.er wizard seguido, el formulario del anterior sigue en el DOM (oculto), y tomar
    el primero que aparece escribía en el equivocado o no encontraba el textarea."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        valor = driver.execute_script(_JS_ESCRIBIR_FORMULA, formula_texto)
        if valor is not None:
            if valor != formula_texto:
                raise RuntimeError("la fórmula escrita en el textarea no coincide con la esperada")
            return
        time.sleep(0.4)
    raise TimeoutError("No se encontró el textarea visible del Formula Builder (isams-textarea#customFormula).")


def guardar_wizard_calculo(driver, tiempo_espera=20):
    """Desde el Step 2 (fórmula ya escrita) hasta el final: Next (-> Step 3)
    -> Next (-> Step 4) -> Next (-> Step 5) -> Finish, para guardar."""
    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)
    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)
    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)
    time.sleep(0.5)
    clic_boton_wizard(driver, "isams-wizard#calculationWizard", "Finish")
    time.sleep(1)


def mapear_columnas(driver, tiempo_espera=60):
    """Pestaña 'Columns' del wizard (ya abierta): recorre la lista UNA sola vez
    (virtualizada, con scroll) y devuelve {nombre_normalizado: [ids]} de TODAS las
    columnas del gradebook, con los ids SIN el prefijo 'ID' (los que van en
    [ID77590]). Misma lógica de contenedor y espera que recolectar_ids_columnas."""
    wizard = elemento_profundo(driver, "isams-wizard#calculationWizard", tiempo_espera)
    contenedor = driver.execute_script(
        "return arguments[0].querySelector("
        "\"markbooks-structure-display[id='columnSelection'] isams-pane[id='listitems']\""
        ");",
        wizard,
    )
    if contenedor is None:
        raise TimeoutError(
            "No se encontró la lista de columnas ('listitems' dentro de "
            "markbooks-structure-display#columnSelection) en el wizard actual."
        )

    fin_espera_inicial = time.time() + 15
    while time.time() < fin_espera_inicial:
        if driver.execute_script(
                "return arguments[0].querySelectorAll(\"div[id^='ID'], div[id^='GID']\").length > 0;", contenedor):
            break
        time.sleep(0.3)

    driver.execute_script("arguments[0].scrollTop = 0;", contenedor)
    time.sleep(0.3)

    mapa = {}
    vistos = set()
    ultimo_top = -1
    sin_cambio = 0
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        filas = driver.execute_script(
            """
            return Array.from(arguments[0].querySelectorAll("div[id^='ID']")).map(function (el) {
                return [el.getAttribute('id'), el.textContent || ''];
            });
            """,
            contenedor,
        )
        for id_attr, texto_completo in filas:
            if not id_attr or id_attr in vistos:
                continue
            vistos.add(id_attr)
            lineas = [l.strip() for l in texto_completo.split("\n") if l.strip()]
            nombre = normalizar(lineas[0]) if lineas else ""
            if nombre:
                mapa.setdefault(nombre, []).append(id_attr[2:])

        info = driver.execute_script(
            """
            const c = arguments[0];
            const puede = c.scrollHeight > c.clientHeight + 5;
            if (puede) { c.scrollTop += c.clientHeight * 0.8; }
            return {top: c.scrollTop, puede: puede};
            """,
            contenedor,
        )
        if not info["puede"]:
            time.sleep(0.3)
            continue
        if info["top"] == ultimo_top:
            sin_cambio += 1
            if sin_cambio >= 3:
                break
        else:
            sin_cambio = 0
        ultimo_top = info["top"]
        time.sleep(0.2)
    return mapa


# --- Lectura de fórmulas ya existentes (solo lectura: se abre el wizard y se cierra SIN guardar) ---

_JS_LEER_FORMULA = """
function recolectar(root, selector, resultados) {
    resultados.push(...root.querySelectorAll(selector));
    for (const nodo of root.querySelectorAll('*')) {
        if (nodo.shadowRoot) recolectar(nodo.shadowRoot, selector, resultados);
    }
}
const candidatos = [];
recolectar(document, 'isams-textarea#customFormula', candidatos);
for (const el of candidatos) {
    if (el.offsetParent === null && el.getClientRects().length === 0) continue;
    const ta = el.shadowRoot ? el.shadowRoot.querySelector('textarea#input') : null;
    if (ta) return ta.value;
}
return null;
"""

_JS_WIZARD_VISIBLE = """
function recolectar(root, selector, resultados) {
    resultados.push(...root.querySelectorAll(selector));
    for (const nodo of root.querySelectorAll('*')) {
        if (nodo.shadowRoot) recolectar(nodo.shadowRoot, selector, resultados);
    }
}
const c = [];
recolectar(document, 'isams-wizard#calculationWizard', c);
return c.some(function (el) { return el.offsetParent !== null || el.getClientRects().length > 0; });
"""


def leer_formula_actual(driver, tiempo_espera=15):
    """Texto de la fórmula en el textarea visible del Formula Builder (Step 2). None si no
    aparece (p. ej. la columna no usa Formula Builder)."""
    fin = time.time() + tiempo_espera
    while time.time() < fin:
        valor = driver.execute_script(_JS_LEER_FORMULA)
        if valor is not None:
            return valor
        time.sleep(0.4)
    return None


def leer_propiedades_paso1(driver, tiempo_espera=10):
    """(tipo_de_calculo, compartir) tal como están en el Step 1, sin modificarlos."""
    try:
        form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)
        textos = []
        for dropdown in ("isams-dropdown#calculation", "isams-dropdown#shareable"):
            host = dentro_de_shadow(driver, form, dropdown, tiempo_espera)
            sel = dentro_de_shadow(driver, host, "select#input", tiempo_espera)
            textos.append(driver.execute_script(
                "const o = arguments[0].options[arguments[0].selectedIndex]; return o ? o.text : '';", sel))
        return textos[0], textos[1]
    except Exception:
        return "", ""


def cerrar_wizard_sin_guardar(driver):
    """Cierra el wizard SIN guardar: botón 'Cancel' si existe; si el wizard sigue visible,
    recarga la página (lo no guardado se descarta). Devuelve True si tuvo que recargar."""
    try:
        clic_boton_wizard(driver, "isams-wizard#calculationWizard", "Cancel", tiempo_espera=3)
    except Exception:
        pass
    time.sleep(1)
    cerrar_alerta_pendiente(driver)  # por si pide confirmar que se descartan cambios
    time.sleep(0.5)
    try:
        sigue = driver.execute_script(_JS_WIZARD_VISIBLE)
    except UnexpectedAlertPresentException:
        cerrar_alerta_pendiente(driver)
        sigue = driver.execute_script(_JS_WIZARD_VISIBLE)
    if sigue:
        driver.refresh()
        time.sleep(3)
        cerrar_alerta_pendiente(driver)
        esperar_filas(driver, 25)
        return True
    return False


# --- Clasificación de materias (ver CLASIFICACION_POR_GRADO en datos.py) ---

_CATEGORIAS_EXACTAS = ("hija", "excluida")


def clasificar_materia(grado, materia):
    """'cuantitativa' (por defecto) | 'cualitativa' | 'hija' | 'madre' | 'manual' | 'excluida'."""
    grado_n = normalizar(grado)
    materia_n = normalizar(materia)
    for prefijos_grado, reglas in CLASIFICACION_POR_GRADO:
        if not any(grado_n.startswith(normalizar(p)) for p in prefijos_grado):
            continue
        for categoria, nombres in reglas.items():
            for nombre in nombres:
                nombre_n = normalizar(nombre)
                if categoria in _CATEGORIAS_EXACTAS:
                    if materia_n == nombre_n:
                        return categoria
                elif materia_n.startswith(nombre_n):
                    return categoria
    return "cuantitativa"


# ---------------------------------------------------------------------------
# Creación de columnas calculadas (menú de la barra lateral -> mismo wizard que al editar)
# ---------------------------------------------------------------------------
_JS_ITEM_MENU_CALCULADA = """
const r = arguments[0].shadowRoot;
if (!r) return null;
const items = Array.from(r.querySelectorAll('div.options-menu-item'));
for (const i of items) { if (/calculated/i.test(i.textContent || '')) return i; }
const grupos = r.querySelectorAll('div.options-menu-group');
if (grupos.length > 1) { const it = grupos[1].querySelectorAll('div.options-menu-item'); if (it.length > 1) return it[1]; }
return null;
"""

_JS_PUBLISHED_NAME = """
// Busca (en el wizard y todos sus shadow roots) un isams-textbox VISIBLE cuyo id contenga 'publish'.
const valor = arguments[1];
function recorrer(raiz) {
  for (const el of raiz.querySelectorAll('*')) {
    if (el.tagName === 'ISAMS-TEXTBOX' && /publish/i.test(el.id || '') && el.offsetParent !== null) {
      const inp = el.shadowRoot && el.shadowRoot.querySelector('input#input');
      if (inp) return inp;
    }
    if (el.shadowRoot) { const r = recorrer(el.shadowRoot); if (r) return r; }
  }
  return null;
}
const inp = recorrer(arguments[0]);
if (!inp) return null;
if (inp.value !== valor) {
  inp.focus(); inp.value = valor;
  inp.dispatchEvent(new Event('input', {bubbles: true, composed: true}));
  inp.dispatchEvent(new Event('change', {bubbles: true, composed: true}));
}
return inp.value;
"""

_JS_REDONDEO = """
// Campo numérico de redondeo (no el dropdown): hosts con id que contenga 'round' dentro del form.
const valor = arguments[1];
const raiz = arguments[0].shadowRoot;
if (!raiz) return null;
for (const el of raiz.querySelectorAll('*')) {
  if (/round/i.test(el.id || '') && el.tagName !== 'ISAMS-DROPDOWN' && el.shadowRoot) {
    const inp = el.shadowRoot.querySelector('input#input');
    if (inp) {
      inp.focus(); inp.value = valor;
      inp.dispatchEvent(new Event('input', {bubbles: true, composed: true}));
      inp.dispatchEvent(new Event('change', {bubbles: true, composed: true}));
      return inp.value;
    }
  }
}
return null;
"""


def abrir_asistente_nueva_columna(driver, tiempo_espera=20):
    """Menú de la barra lateral (isams-options-menu#sidebarConfig) -> 'Add Calculated Column'."""
    menu = elemento_profundo(driver, "isams-options-menu#sidebarConfig", tiempo_espera)
    fin = time.time() + tiempo_espera
    clic_en_menu = False
    while time.time() < fin:
        item = driver.execute_script(_JS_ITEM_MENU_CALCULADA, menu)
        if item is not None and driver.execute_script("return arguments[0].offsetParent !== null;", item):
            clic_seguro(driver, item)
            time.sleep(1.5)
            elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)
            return
        if not clic_en_menu:  # el menú puede estar cerrado: lo abrimos clicando el propio componente
            clic_seguro(driver, menu)
            clic_en_menu = True
        time.sleep(0.4)
    raise TimeoutError("No se pudo abrir 'Add Calculated Column' en el menú (isams-options-menu#sidebarConfig).")


def poner_texto_form(driver, form, id_host, valor, tiempo_espera=20):
    host = dentro_de_shadow(driver, form, f"isams-textbox#{id_host}", tiempo_espera)
    inp = dentro_de_shadow(driver, host, "input#input", tiempo_espera)
    driver.execute_script(
        "const i=arguments[0]; i.focus(); i.value=arguments[1];"
        "i.dispatchEvent(new Event('input',{bubbles:true,composed:true}));"
        "i.dispatchEvent(new Event('change',{bubbles:true,composed:true}));", inp, valor)
    if (inp.get_attribute("value") or "") != valor:
        raise RuntimeError(f"no se pudo escribir '{valor}' en {id_host}")


def _elegir_en_dropdown(driver, form, id_host, texto, tiempo_espera=20):
    host = dentro_de_shadow(driver, form, f"isams-dropdown#{id_host}", tiempo_espera)
    sel = dentro_de_shadow(driver, host, "select#input", tiempo_espera)
    objetivo = normalizar(texto)
    for op in sel.find_elements(By.TAG_NAME, "option"):
        if normalizar(op.get_attribute("textContent") or "") == objetivo:
            Select(sel).select_by_visible_text(op.text)
            return
    raise RuntimeError(f"'{texto}' no es una opción de {id_host}")


def rellenar_propiedades_nueva_columna(driver, nombre, nombre_corto, carpeta, redondeo_tipo, redondeo_valor,
                                       tiempo_espera=20):
    """Paso 1 del asistente de creación: nombre, short name, carpeta (Position), Formula Builder,
    redondeo y Allow Sharing (COMPARTIR_CALCULADAS). Termina con Next."""
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)
    poner_texto_form(driver, form, "columnName", nombre, tiempo_espera)
    poner_texto_form(driver, form, "columnShortName", nombre_corto, tiempo_espera)
    _elegir_en_dropdown(driver, form, "position", carpeta, tiempo_espera)
    _elegir_en_dropdown(driver, form, "calculation", "Formula Builder", tiempo_espera)
    time.sleep(0.5)
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)  # puede re-renderizar
    _elegir_en_dropdown(driver, form, "calculationRoundingType", redondeo_tipo, tiempo_espera)
    time.sleep(0.5)
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)
    valor = driver.execute_script(_JS_REDONDEO, form, redondeo_valor)
    if valor is None:
        print("   (aviso: no se encontró el campo numérico de Rounding; se deja el valor por defecto)")
    _elegir_en_dropdown(driver, form, "shareable", COMPARTIR_CALCULADAS, tiempo_espera)
    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)


def terminar_asistente_creacion(driver, nombre_publicado, tiempo_espera=20):
    """Desde el paso de la fórmula (ya escrita): avanza con Next, rellenando el Published Name con el
    Short Name en el paso que lo tenga, hasta que aparezca 'Finish', y lo clica."""
    for _ in range(8):
        time.sleep(0.5)
        wizard = elemento_profundo(driver, "isams-wizard#calculationWizard", tiempo_espera)
        driver.execute_script(_JS_PUBLISHED_NAME, wizard, nombre_publicado)
        try:
            clic_boton_wizard(driver, "isams-wizard#calculationWizard", "Finish", tiempo_espera=2)
            time.sleep(1.5)
            return
        except TimeoutError:
            clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard", tiempo_espera)
    raise TimeoutError("No apareció el botón 'Finish' del asistente de creación.")