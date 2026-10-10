
"""
Configuración centralizada del proceso de automatización de iSAMS.
Ajusta estos valores según tu entorno; el script principal solo importa
de aquí, no debe tener rutas ni listas "quemadas" adentro.
"""
 
# --- Navegador ---
CHROME_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chrome-win64\chrome.exe"
DRIVER_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chromedriver-win64\chromedriver.exe"
# --- iSAMS ---
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
# Filas cuyo nombre empiece por estos prefijos NUNCA se toman como insumos (aunque lleven %).
PREFIJOS_NO_INSUMO = ("SUMAINST", "CALINST", "CALEXAT", "EQUI", "PROMTRI")
# Si la fila calculada no trae el % en su Short Name, se usan estos valores.
PORC_INSUMOS_DEFECTO = 70
PORC_EXAMEN_DEFECTO = 30
 
# FLOOR(valor, 0.01) trunca (no redondea) a centésimas, igual que las fórmulas actuales del colegio.
FLOOR_SIGNIFICANCIA = "0.01"
 
ARCHIVO_FORMULAS = "formulas_{promocion}.xlsx"
 
ARCHIVO_RESULTADO = "resultado_{promocion}.xlsx"   # estado (ok/error) de cada fórmula escrita
 