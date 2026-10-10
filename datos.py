
"""
Configuración centralizada del proceso de automatización de iSAMS.
Ajusta estos valores según tu entorno; el script principal solo importa
de aquí, no debe tener rutas ni listas "quemadas" adentro.
"""
 
# --- Navegador ---
CHROME_PATH = r"C:\Users\egaib\Desktop\Documentos Evans\Documentos Evans\Scripts\chrome-win64\chrome.exe" 
DRIVER_PATH = r"C:\Users\egaib\Desktop\Documentos Evans\Documentos Evans\Scripts\chromedriver-win64\chromedriver.exe" # ej: r"C:\ruta\chromedriver.exe" (None = Selenium Manager lo resuelve solo)
 
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