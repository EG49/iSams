"""
Automatización iSAMS: arma las fórmulas de "Calificaciones Finales" (PROM
General Q1P1, Q1P2, Q1, Q2P1, Q2P2, Q2, FINAL, y APROBACION) promediando
todas las columnas compartidas correspondientes de cada materia.

Reutiliza la configuración de datos.py y las funciones de bajo nivel ya
probadas en isams_promedios.py (login, navegación, manejo de shadow DOM,
scroll en el árbol virtualizado, etc.) en vez de duplicarlas.

Estado actual: Paso 1 (abrir gradebook + ubicar carpeta) y Paso 2 (armar y
escribir la fórmula de "PROM General Q1P1" como prueba puntual) ya
funcionan. Falta confirmar los pasos 3, 4 y 5 del wizard "Edit Calculated
Column" (dice "Step 1 of 5"; el 1 es propiedades, el 2 es la fórmula) antes
de poder guardar y automatizar el resto de filas ("PROM General" X8, más
APROBACION) y el loop completo por grado.
"""

import time
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

from datos import (
    CHROME_PATH,
    DRIVER_PATH,
    LOGIN_URL,
    EXCEL_PATH,
    HOJA_GRADES,
)

from isams_promedios import (
    cargar_datos_excel,
    ir_a_lista_gradebooks,
    buscar_fila_por_nombre,
    buscar_id_carpeta,
    clic_seguro,
    elemento_profundo,
    dentro_de_shadow,
    buscar_elementos_con_scroll,
    xpath_literal,
    buscar_boton_editar_columna,
    clic_next_wizard_verificado,
    clic_boton_wizard,
    cerrar_alerta_pendiente,
    normalizar,
)


NOMBRE_CARPETA_FINALES = "Calificaciones Finales"

# Mapeo confirmado: fila "PROM General {sigla}" -> nombre de la columna
# fuente a promediar entre todas las materias del gradebook.
MAPEO_PROM_GENERAL = {
    "Q1P1": "TOTAL Q1 PARCIAL 1",
    "Q1P2": "TOTAL Q1 PARCIAL 2",
    "Q1": "PROMEDIO QUIMESTRE 1",
    "Q2P1": "TOTAL Q2 PARCIAL 1",
    "Q2P2": "TOTAL Q2 PARCIAL 2",
    "Q2": "PROMEDIO QUIMESTRE 2",
    "FINAL": "PROMEDIO FINAL",
}

# Caso aparte: normalmente promedia "PROMEDIO FINAL SUPLETORIO" de cada
# materia como las demás, EXCEPTO en 2do/3ro/4to EGB, donde en vez de eso
# es una referencia directa al id de "PROM General FINAL" (ver más abajo).
NOMBRE_FILA_FINAL_SUPLETORIO = ["PROM General FINAL SUPLETORIO", "Promedio General Final Supletorio"]
NOMBRE_FILA_FINAL_SUPLETORIO_TXT = " / ".join(NOMBRE_FILA_FINAL_SUPLETORIO)
COLUMNA_FUENTE_FINAL_SUPLETORIO = "PROMEDIO FINAL SUPLETORIO"

# Grados que además llevan las 7 filas "... Cualitativo" y el cambio de
# FINAL SUPLETORIO a referencia directa.
PREFIJOS_CON_CUALITATIVO = ("2EGB", "3EGB", "4EGB")


def es_grado_con_cualitativo(grado):
    """True si el grado empieza por 2EGB, 3EGB o 4EGB (ej. '2EGB-A',
    '3EGB-B', '4EGB-C'), insensible a mayúsculas/tildes/espacios."""
    grado_norm = normalizar(grado)
    return any(
        re.match(r"^" + re.escape(p) + r"(?=[\s\-]|$)", grado_norm)
        for p in PREFIJOS_CON_CUALITATIVO
    )


# --- Paso 1: ubicar y abrir el gradebook / la carpeta ---

def buscar_fila_por_titulo_con_scroll(driver, nombre_fila, tiempo_espera=45):
    """Ubica una fila del árbol de estructura (principal, el de las
    materias) por su título exacto, scrolleando si hace falta. nombre_fila
    puede ser un string o una lista de nombres alternativos (la misma fila
    a veces se llama distinto según el grado, ej. 'PROM General FINAL
    SUPLETORIO' vs 'Promedio General Final Supletorio')."""
    nombres = [nombre_fila] if isinstance(nombre_fila, str) else list(nombre_fila)
    condicion = " or ".join(f"@title={xpath_literal(n)}" for n in nombres)
    filas = buscar_elementos_con_scroll(
        driver,
        f"//div[contains(@class,'row')][{condicion}]",
        tiempo_espera=tiempo_espera,
    )
    if not filas:
        nombres_str = " / ".join(repr(n) for n in nombres)
        raise TimeoutError(f"No se encontró ninguna fila con título {nombres_str} en el árbol de estructura.")
    return filas[0]


def abrir_gradebook_y_ubicar_calificaciones_finales(driver, wait, grado):
    """Abre el gradebook 'Promedios {grado}' y localiza (scrolleando si
    hace falta) la carpeta 'Calificaciones Finales'. Devuelve su id."""
    ir_a_lista_gradebooks(driver, wait)

    nombre_fila = f"Promedios {grado}"
    fila_objetivo = buscar_fila_por_nombre(driver, nombre_fila)
    if fila_objetivo is None:
        raise RuntimeError(f"gradebook '{nombre_fila}' no encontrado en la lista")

    icono = fila_objetivo.find_element(By.CSS_SELECTOR, "isams-icon[title='View Gradebook']")
    clic_seguro(driver, icono)

    WebDriverWait(driver, 20).until(lambda d: re.search(r"/markbooks/\w+/(\d+)", d.current_url))
    match = re.search(r"/markbooks/\w+/(\d+)", driver.current_url)
    gradebook_id = match.group(1)

    edit_url = LOGIN_URL.rstrip("/") + f"/main/areas/markbooks/index.html#/markbooks/edit/{gradebook_id}"
    driver.get(edit_url)
    time.sleep(2)

    carpeta_id = buscar_id_carpeta(driver, NOMBRE_CARPETA_FINALES)
    return carpeta_id


# --- Paso 2: wizard "Edit Calculated Column" ---

def configurar_propiedades_calculo(driver, tiempo_espera=20):
    """Step 1 del wizard: Calculation Type = 'Formula Builder',
    Allow Sharing = 'Shared as Read Only', y avanza al Step 2."""
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", tiempo_espera)

    dropdown_calculo = dentro_de_shadow(driver, form, "isams-dropdown#calculation", tiempo_espera)
    select_calculo = dentro_de_shadow(driver, dropdown_calculo, "select#input", tiempo_espera)
    Select(select_calculo).select_by_visible_text("Formula Builder")

    dropdown_compartir = dentro_de_shadow(driver, form, "isams-dropdown#shareable", tiempo_espera)
    select_compartir = dentro_de_shadow(driver, dropdown_compartir, "select#input", tiempo_espera)
    Select(select_compartir).select_by_visible_text("Not Shared")

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


def construir_formula_promedio(ids):
    """Arma la fórmula anidada if(ISBLANK(...),"",...) + FLOOR(AVERAGE(...))
    a partir de una lista de ids de columna (sin el prefijo 'ID'), en el
    mismo formato que usa iSAMS."""
    if not ids:
        raise ValueError("No hay ids para construir la fórmula.")
    refs = [f"[ID{i}]" for i in ids]
    formula = ""
    for ref in refs:
        formula += f'if(ISBLANK({ref},true,false),"",\n'
    formula += f'FLOOR(AVERAGE({",".join(refs)},"RAW"),0.01)'
    formula += ")" * len(refs)
    return formula


def escribir_formula(driver, formula_texto, tiempo_espera=20):
    """Escribe (reemplazando el contenido) la fórmula en el textarea del
    Formula Builder, disparando los eventos que la app espera."""
    form = elemento_profundo(driver, "isams-data-form#customFormulaForm", tiempo_espera)
    contenedor_textarea = dentro_de_shadow(driver, form, "isams-textarea#customFormula", tiempo_espera)
    textarea = dentro_de_shadow(driver, contenedor_textarea, "textarea#input", tiempo_espera)
    driver.execute_script(
        """
        const ta = arguments[0];
        ta.value = arguments[1];
        ta.dispatchEvent(new Event('input', {bubbles: true}));
        ta.dispatchEvent(new Event('change', {bubbles: true}));
        """,
        textarea, formula_texto,
    )


def armar_y_escribir_formula_promedio(driver, nombre_columna_fuente, tiempo_espera=60):
    """Desde el Step 2 del wizard ya abierto: clica 'Columns', recolecta
    los ids de nombre_columna_fuente (con scroll), arma la fórmula y la
    escribe en el Formula Builder. Devuelve (formula, ids) para poder
    inspeccionar/loggear."""
    abrir_pestana_columns(driver, tiempo_espera)
    ids = recolectar_ids_columnas(driver, nombre_columna_fuente, tiempo_espera)
    if not ids:
        raise RuntimeError(f"No se encontraron columnas '{nombre_columna_fuente}' en la pestaña Columns.")
    formula = construir_formula_promedio(ids)
    escribir_formula(driver, formula, tiempo_espera)
    return formula, ids


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


def construir_formula_aprobacion(id_final):
    """Arma la fórmula de APROBACION: IF([ID del 'PROM General FINAL' de
    ESTE gradebook] >= "7", "APROBADO", "REPROBADO"). A diferencia de las
    otras 7 filas, depende de un ÚNICO id (no promedia varias materias)."""
    return f'IF([ID{id_final}]>="7","APROBADO","REPROBADO")'


def construir_formula_referencia_directa(id_valor):
    """Fórmula que simplemente copia el valor de otra fila del mismo
    gradebook (ej. 'PROM General FINAL SUPLETORIO' = 'PROM General FINAL'
    en 2do/3ro/4to EGB)."""
    return f"[ID{id_valor}]"


def construir_formula_cualitativo(id_valor):
    """Fórmula de escala cualitativa (E- a A+) basada en el id numérico de
    la fila 'PROM General X' hermana dentro del mismo gradebook. Se arma
    de adentro hacia afuera para no equivocarse contando paréntesis a mano."""
    ref = f"[ID{id_valor}]"
    escalones = [
        (0.1, " "),
        (1.49, "E-"),
        (2.49, "E+"),
        (3.49, "D-"),
        (4.49, "D+"),
        (5.49, "C-"),
        (6.49, "C+"),
        (7.49, "B-"),
        (8.49, "B+"),
        (9.49, "A-"),
    ]
    interior = f'IF({ref}<=10,"A+"," ")'
    for umbral, etiqueta in reversed(escalones):
        interior = f'IF({ref}<{umbral},"{etiqueta}",{interior})'
    return f'IF(ISBLANK({ref},true,false),"",{interior})'


def obtener_id_unico(driver, nombre_columna, tiempo_espera=60):
    """Dentro de la pestaña 'Columns' (ya abierta), devuelve el ÚNICO id
    que coincide con nombre_columna. Falla si hay 0 o más de 1 coincidencia
    (para estas fórmulas, se espera una sola: la fila hermana dentro del
    mismo gradebook, no repetida por materia)."""
    ids = recolectar_ids_columnas(driver, nombre_columna, tiempo_espera)
    if not ids:
        raise RuntimeError(f"No se encontró '{nombre_columna}' en la pestaña Columns.")
    if len(ids) > 1:
        raise RuntimeError(f"Se esperaba un único '{nombre_columna}', se encontraron {len(ids)}: {ids}")
    return ids[0]


def procesar_fila_por_id_unico(driver, nombre_fila, nombre_fila_fuente, construir_formula_fn, tiempo_espera=60):
    """Flujo genérico para una fila cuya fórmula depende de un ÚNICO id (el
    de nombre_fila_fuente, dentro del mismo gradebook, sin promediar varias
    materias): ubicar la fila, abrir su editor, Step 1, pestaña Columns,
    obtener ese id único, construir la fórmula con construir_formula_fn(id),
    escribirla y guardar. Reutilizado por APROBACION, la referencia directa
    de FINAL SUPLETORIO, y las filas Cualitativo."""
    fila = buscar_fila_por_titulo_con_scroll(driver, nombre_fila, tiempo_espera)
    boton_editar = buscar_boton_editar_columna(driver, fila, tiempo_espera)
    clic_seguro(driver, boton_editar)
    time.sleep(1.5)

    configurar_propiedades_calculo(driver, tiempo_espera)

    abrir_pestana_columns(driver, tiempo_espera)
    id_valor = obtener_id_unico(driver, nombre_fila_fuente, tiempo_espera)
    formula = construir_formula_fn(id_valor)
    escribir_formula(driver, formula, tiempo_espera)
    guardar_wizard_calculo(driver, tiempo_espera)
    return formula, id_valor


def procesar_fila_prom_general(driver, nombre_fila, nombre_columna_fuente, tiempo_espera=60):
    """Flujo completo para UNA fila de promedio general (ej. 'PROM General
    Q1P1' o 'Promedio General Final Supletorio'): ubicarla, abrir su
    editor, configurar Step 1, armar y escribir la fórmula (promedio de
    todas las 'nombre_columna_fuente' del gradebook), y guardar. Lanza la
    excepción tal cual si algo falla (el llamador decide cómo reportarlo)."""
    fila = buscar_fila_por_titulo_con_scroll(driver, nombre_fila, tiempo_espera)
    boton_editar = buscar_boton_editar_columna(driver, fila, tiempo_espera)
    clic_seguro(driver, boton_editar)
    time.sleep(1.5)

    configurar_propiedades_calculo(driver, tiempo_espera)

    formula, ids = armar_y_escribir_formula_promedio(driver, nombre_columna_fuente, tiempo_espera)
    guardar_wizard_calculo(driver, tiempo_espera)
    return formula, ids


def main():
    print(f"Leyendo '{EXCEL_PATH}'...")
    grados, _codigos_no_usados = cargar_datos_excel()
    print(f"{len(grados)} grados cargados de la hoja '{HOJA_GRADES}'.")

    options = webdriver.ChromeOptions()
    if CHROME_PATH:
        options.binary_location = CHROME_PATH

    if DRIVER_PATH:
        service = Service(DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)

    driver.maximize_window()
    driver.get(LOGIN_URL)

    input("\n>>> Inicia sesión manualmente en la ventana del navegador.\n"
          ">>> Cuando ya estés dentro de la plataforma, vuelve aquí y presiona ENTER...\n")

    wait = WebDriverWait(driver, 20)
    reporte_faltantes = []

    for grado in grados:
        print(f"\n=== {grado} ===")
        try:
            carpeta_id = abrir_gradebook_y_ubicar_calificaciones_finales(driver, wait, grado)
            print(f"✅ Carpeta '{NOMBRE_CARPETA_FINALES}' localizada (id={carpeta_id}).")
        except Exception as e:
            mensaje = f"{grado}: ERROR al abrir el gradebook - {e}"
            print(f"⚠️  {mensaje}")
            reporte_faltantes.append(mensaje)
            cerrar_alerta_pendiente(driver)
            continue  # sin gradebook abierto no tiene caso seguir con este grado

        for sigla, nombre_columna_fuente in MAPEO_PROM_GENERAL.items():
            nombre_fila = f"PROM General {sigla}"
            print(f"\n-- {nombre_fila} (fuente: '{nombre_columna_fuente}') --")
            try:
                formula, ids = procesar_fila_prom_general(driver, nombre_fila, nombre_columna_fuente)
                print(f"✅ '{nombre_fila}' guardada, usando {len(ids)} columna(s): {ids}")
            except Exception as e:
                mensaje = f"{grado} / {nombre_fila}: ERROR - {e}"
                print(f"⚠️  {mensaje}")
                reporte_faltantes.append(mensaje)
                cerrar_alerta_pendiente(driver)

        # PROM General FINAL SUPLETORIO: promedio normal, salvo en
        # 2do/3ro/4to EGB, donde es una referencia directa al id de
        # "PROM General FINAL Cualitativo" (no promedia varias materias).
        print(f"\n-- {NOMBRE_FILA_FINAL_SUPLETORIO_TXT} --")
        try:
            if es_grado_con_cualitativo(grado):
                formula, id_valor = procesar_fila_por_id_unico(
                    driver, NOMBRE_FILA_FINAL_SUPLETORIO, "PROM General FINAL Cualitativo",
                    construir_formula_referencia_directa,
                )
                print(f"✅ '{NOMBRE_FILA_FINAL_SUPLETORIO_TXT}' guardada como referencia directa "
                      f"a PROM General FINAL Cualitativo (id={id_valor}).")
            else:
                formula, ids = procesar_fila_prom_general(
                    driver, NOMBRE_FILA_FINAL_SUPLETORIO, COLUMNA_FUENTE_FINAL_SUPLETORIO
                )
                print(f"✅ '{NOMBRE_FILA_FINAL_SUPLETORIO_TXT}' guardada, usando {len(ids)} columna(s): {ids}")
        except Exception as e:
            mensaje = f"{grado} / {NOMBRE_FILA_FINAL_SUPLETORIO_TXT}: ERROR - {e}"
            print(f"⚠️  {mensaje}")
            reporte_faltantes.append(mensaje)
            cerrar_alerta_pendiente(driver)

        print("\n-- APROBACION --")
        try:
            formula, id_final = procesar_fila_por_id_unico(
                driver, "APROBACION", "PROM General FINAL", construir_formula_aprobacion
            )
            print(f"✅ 'APROBACION' guardada, usando PROM General FINAL (id={id_final}).")
        except Exception as e:
            mensaje = f"{grado} / APROBACION: ERROR - {e}"
            print(f"⚠️  {mensaje}")
            reporte_faltantes.append(mensaje)
            cerrar_alerta_pendiente(driver)

        # Filas "... Cualitativo": solo para 2do/3ro/4to EGB. Cada una
        # referencia el id único de su fila numérica hermana (ej.
        # "PROM General Q1P1 Cualitativo" usa el id de "PROM General Q1P1").
        if es_grado_con_cualitativo(grado):
            print(f"\n-- Filas Cualitativo ({grado}) --")
            for sigla in MAPEO_PROM_GENERAL.keys():
                nombre_fila_num = f"PROM General {sigla}"
                nombre_fila_cual = f"{nombre_fila_num} Cualitativo"
                print(f"\n-- {nombre_fila_cual} --")
                try:
                    formula, id_valor = procesar_fila_por_id_unico(
                        driver, nombre_fila_cual, nombre_fila_num, construir_formula_cualitativo
                    )
                    print(f"✅ '{nombre_fila_cual}' guardada (basada en id={id_valor}).")
                except Exception as e:
                    mensaje = f"{grado} / {nombre_fila_cual}: ERROR - {e}"
                    print(f"⚠️  {mensaje}")
                    reporte_faltantes.append(mensaje)
                    cerrar_alerta_pendiente(driver)

    print("\n===== RESUMEN =====")
    if reporte_faltantes:
        print(f"Elementos con problemas ({len(reporte_faltantes)}):")
        for item in reporte_faltantes:
            print(f"  - {item}")
    else:
        print("Prueba completada sin errores. ✅")

    input("\nPresiona ENTER para cerrar el navegador...")
    driver.quit()


if __name__ == "__main__":
    main()