
"""
Configuración centralizada del proceso de automatización de iSAMS.
Ajusta estos valores según tu entorno; el script principal solo importa
de aquí, no debe tener rutas ni listas "quemadas" adentro.
"""
 
# --- Navegador ---
CHROME_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chrome-win64\chrome.exe"
DRIVER_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chromedriver-win64\chromedriver.exe"

# --- iSAMS ---
LOGIN_URL = "https://lpcentenario.isamshosting.cloud/"
 
# Ruta del editor de un gradebook; {id} se reemplaza por el ID de la materia.
# Página del módulo Gradebooks (se carga UNA vez; luego solo cambia el #hash, como una SPA)
RUTA_MODULO_GRADEBOOKS = "main/areas/markbooks/index.html"
RUTA_EDITAR_GRADEBOOK = "main/areas/markbooks/index.html#/markbooks/edit/{id}"
 
# --- Datos de entrada ---
EXCEL_PATH = "gradebooks.xlsx"
HOJA_CODES = "codes"                # fila 1 = encabezado; col A = código (ID), col B = Materia (nombre formal, ej. "3BCH-A FÍSICA")
PROMOCION_POR_DEFECTO = "3BCH"      # o pásala por consola: python leer_materias.py 3BCH
 
# Materias a saltar (texto exacto de la columna 'materia', sin importar mayúsculas/tildes).
EXCLUIR_MATERIAS = []               # ej. ["PROMEDIO", "COMPORTAMIENTO", "CAS"]
 
# Carpeta (trimestre) a leer dentro de cada materia; se compara sin importar mayúsculas/tildes.
CARPETA_A_LEER = "TII"
 
# Nombres alternos de la carpeta (algunas materias las llamaron "TRIMESTRE I" en vez de "TI").
# Se acepta el nombre de CARPETA_A_LEER o cualquiera de sus alias (sin importar mayúsculas/tildes).
ALIAS_CARPETAS = {"TI": ["TRIMESTRE I"], "TII": ["TRIMESTRE II"], "TIII": ["TRIMESTRE III"]}
 
# --- Salida ---
# {promocion} se reemplaza, ej. lectura_3BCH.xlsx
ARCHIVO_SALIDA = "lectura_{promocion}.xlsx"
 
# --- Fórmulas del trimestre (paso 2) ---
# Nombres de las filas calculadas: se detectan por el INICIO del nombre (Name) de la fila.
PREFIJO_SUMA_INSUMOS = "SUMAINST"     # Sumatoria Insumo 70% T-I
PREFIJO_CAL_INSUMOS = "CALINST"       # Calculo / PROM Insumo 70%
PREFIJO_EXAMEN = "EXAT"               # Examen T-I (nota plana, no se toca)
PREFIJO_CAL_EXAMEN = "CALEXAT"        # Calculo Examen 30% T-I
PREFIJO_PROMEDIO_TRIMESTRE = "PROMTRI"  # PROMEDIO TRIMESTRE I
# --- Materias cualitativas (un solo insumo): INSCUALT1 + sus filas de equivalencia ---
PREFIJO_INSUMO_CUALITATIVO = "INSCUALT"   # Insumo Cualitativo T-I (nota de entrada, no se toca)
# (prefijo de la fila de equivalencia, prefijo de la fila de la que toma el valor). Las equivalencias usan
# la escala A+..E- (la anterior A/P/I está deprecada). Solo se escriben las filas que existan.
EQUIVALENCIAS_CUALITATIVA = [
    ("EQUINSUT", "INSCUALT"),         # Equivalencia del Insumo T-I
    ("EQUICALNSUT", "CALINST"),       # Equivalencia del Cálculo 70% T-I
    ("EQUIEXAT", "EXAT"),             # Equivalencia de Examen Trimestre I
    ("EQUICALEXAT", "CALEXAT"),       # Equivalencia Calculo Examen 30% T-I
    ("EQUICALPROMTR", "PROMTRI"),     # Equivalencia Promedio T-I
]
 
# Nombres alternos de las columnas calculadas (inicio del Name), p. ej. las de 5EGB-7EGB que usan "H2":
#   INSH2T2 = Sumatoria Insumo, CALINSH2T2 = Calculo Insumo, EXAH2T2 / EXA2 = Examen, PROMH2TRI2 = Promedio trimestre.
# Se acepta el prefijo estándar o cualquiera de sus alias.
ALIAS_PREFIJOS = {
    "SUMAINST": ["INSH"],
    "CALINST": ["CALINSH"],
    "EXAT": ["EXAH", "EXA"],
    "PROMTRI": ["PROMH"],
}
 
# Equivalencia en letras del promedio del trimestre (EQUPROT1 = "Equi Promedio T-I"). Solo se genera si la fila existe.
PREFIJO_EQUIV_PROMEDIO = "EQUPRO"
# Filas cuyo nombre empiece por estos prefijos NUNCA se toman como insumos (aunque lleven %).
PREFIJOS_NO_INSUMO = ("SUMAINST", "CALINST", "CALEXAT", "EQU", "PROMTRI")
# Si la fila calculada no trae el % en su Short Name, se usan estos valores.
PORC_INSUMOS_DEFECTO = 70
PORC_EXAMEN_DEFECTO = 30
 
# FLOOR(valor, 0.01) trunca (no redondea) a centésimas, igual que las fórmulas actuales del colegio.
FLOOR_SIGNIFICANCIA = "0.01"
 
ARCHIVO_FORMULAS = "formulas_{promocion}.xlsx"
 
ARCHIVO_RESULTADO = "resultado_{promocion}.xlsx"   # estado (ok/error) de cada fórmula escrita
ARCHIVO_FORMULAS_LEIDAS = "formulas_leidas_{etiqueta}.xlsx"   # fórmulas que YA existen en iSAMS (guía)
 
# Valor de "Allow Sharing" (Step 1 del wizard) al escribir fórmulas. Igual que 4EGB-A, que está bien hecho.
COMPARTIR_CALCULADAS = "Shared as Read Only"
 
# --- Clasificación de materias por grado ---
# Las que NO son "cuantitativa" ni "cualitativa" no entran al paso de fórmulas: van a la hoja 'omitidas' de
# formulas_{promocion}.xlsx con su categoría, para tratarlas con su propio proceso.
#   cuantitativa = el caso normal (SUMAINST/CALINST/CALEXAT/PROMTRI [+ EQUPRO])   <- por defecto
#   cualitativa  = un solo insumo (INSCUALT): fórmulas propias, ver formulas_cualitativa
#   cualitativa_mh = cualitativas de 5EGB-7EGB hechas por otra persona (12 columnas por trimestre, "M-H"):
#                    se identifican por POSICIÓN en la carpeta, ver formulas_cualitativa_mh
#   hija         = Art / Música, hijas de Educación Cultural y Artística; proceso aparte (pendiente)
#   madre        = Educación Cultural y Artística (promedio de sus hijas); proceso aparte (pendiente)
#   manual       = COMPORTAMIENTO: se llena a mano, no se toca
#   excluida     = no es una materia de este proceso (el gradebook 'PROMEDIO' va por el otro proyecto)
# Cada lista son INICIOS de nombre, comparados sin mayúsculas/tildes. Un grado que no esté
# en ninguna regla (ej. 3BCH) se trata todo como cuantitativa.
CLASIFICACION_POR_GRADO = [
    (("2EGB", "3EGB", "4EGB"), {
        "cualitativa": [
            "ACOMPAÑAMIENTO INTEGRAL EN EL AULA",
            "DIGITAL LEARNING",
            "EDUCACIÓN FÍSICA",
            "EDUCACIÓN PARA LA VIDA/FORMACIÓN CRISTIANA",
            "ROBOTICS",
        ],
        "hija": ["ART", "MÚSICA", "MUSIC"],          # coincidencia EXACTA (no inicio)
        "madre": ["EDUCACIÓN CULTURAL Y ARTÍSTICA"],
        "manual": ["COMPORTAMIENTO"],
        "excluida": ["PROMEDIO"],                     # coincidencia EXACTA
    }),
    (("5EGB", "6EGB", "7EGB"), {
        "cualitativa_mh": [
            "DIGITAL LEARNING",
            "EDUCACIÓN PARA LA VIDA/FORMACIÓN CRISTIANA",
            "ROBOTICS",
            "ACOMPAÑAMIENTO INTEGRAL EN EL AULA",     # cubre ".../HABLEMOS DE APRENDIZAJE"
        ],
        "manual": ["COMPORTAMIENTO"],
        "excluida": ["PROMEDIO"],
    }),
]
 
 
# --- Creación de columnas calculadas (crear_columnas.py) ---
# Trimestre según la carpeta: nombre numérico (SUMAINST2) y romano en el Short Name (T-II).
TRIMESTRES = {"TI": (1, "I"), "TII": (2, "II"), "TIII": (3, "III")}
REDONDEO_TIPO = "Rounding to a given decimal"
REDONDEO_DECIMALES = "3"
ARCHIVO_PLAN_COLUMNAS = "plan_columnas_{promocion}.xlsx"
ARCHIVO_RESULTADO_CREACION = "creacion_{promocion}.xlsx"
# Columnas calculadas por tipo, EN EL ORDEN DE CREACIÓN (cada una solo depende de las anteriores
# o de INSCUALT/EXAT, que ya existen): (prefijo para detectar si existe, nombre base, Short Name).
# En el nombre se añade el número del trimestre; en el Short Name {R} es el romano (I, II, III).
COLUMNAS_CUANTITATIVA = [
    ("SUMAINST", "SUMAINST", "Sumatoria Insumo 70% T-{R}"),
    ("CALINST", "CALINST", "Calculo Insumo 70% T-{R}"),
    ("CALEXAT", "CALEXAT", "Calculo Examen 30% T-{R}"),
    ("PROMTRI", "PROMTRI", "PROMEDIO TRIMESTRE {R}"),
    ("EQUPRO", "EQUPROT", "Equi Promedio T-{R}"),
]
COLUMNAS_CUALITATIVA = [
    ("EQUINSUT", "EQUINSUT", "Equivalencia del Insumo T-{R}"),
    ("CALINST", "CALINST", "Calculo Insumo 70% T-{R}"),
    ("EQUICALNSUT", "EQUICALNSUT", "Equivalencia del Cálculo 70% T-{R}"),
    ("EQUIEXAT", "EQUIEXAT", "Equivalencia de Examen Trimestre {R}"),
    ("CALEXAT", "CALEXAT", "Calculo Examen 30% T-{R}"),
    ("EQUICALEXAT", "EQUICALEXAT", "Equivalencia Calculo Examen 30% T-{R}"),
    ("PROMTRI", "PROMTRI", "PROMEDIO TRIMESTRE {R}"),
    ("EQUICALPROMTR", "EQUICALPROMTR", "Equivalencia Promedio T-{R}"),
]
