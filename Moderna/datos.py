
"""
Configuración centralizada del proceso de automatización de iSAMS.
Ajusta estos valores según tu entorno; el script principal solo importa
de aquí, no debe tener rutas ni listas "quemadas" adentro.
"""
 
# --- Navegador ---
CHROME_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chrome-win64\chrome.exe"
DRIVER_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chromedriver-win64\chromedriver.exe"
# --- iSAMS ---
LOGIN_URL = "https://lamoderna.isamshosting.cloud/"
 
# --- Excel ---
EXCEL_PATH = "gradebooks.xlsx"
HOJA_GRADES = "Grades"   # columna A
HOJA_CODES = "Codes"     # columnas: Codigo
 
# --- Columnas a marcar
COLUMNAS_A_MARCAR = [
    "TOTAL Q1 PARCIAL 1",
    "TOTAL Q1 PARCIAL 2",
    "PROMEDIO QUIMESTRE 1",
    "TOTAL Q2 PARCIAL 1",
    "TOTAL Q2 PARCIAL 2",
    "PROMEDIO QUIMESTRE 2",
    "PROMEDIO FINAL",
    "PROMEDIO FINAL SUPLETORIO",
]