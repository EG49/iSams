
"""
Configuración centralizada del proceso de automatización de iSAMS.
Ajusta estos valores según tu entorno; el script principal solo importa
de aquí, no debe tener rutas ni listas "quemadas" adentro.
"""
 
# --- Navegador ---
CHROME_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chrome-win64\chrome.exe"
DRIVER_PATH = r"C:\Users\egaib\Desktop\Programación\iSams\chromedriver-win64\chromedriver.exe"
# --- iSAMS ---
LOGIN_URL = "https://lpcentenario.isamshosting.cloud/auth/Account/Login?ReturnUrl=%2Fauth%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3DiSAMS%26redirect_uri%3Dhttps%253A%252F%252Flpcentenario.isamshosting.cloud%252Fmain%252Fframework%26response_type%3Dcode%2520id_token%2520token%26scope%3Dopenid%2520isamsidentity%2520restapi%2520authapi%2520offline_access%2520iSAMS.Auth.Client.Create%2520iSAMS.Auth.Client.Delete%26state%3DOpenIdConnect.AuthenticationProperties%253DwnnUvRRNxYLz0O8xQGH9_qBNb3dIUjBm3Te4t3ZqgGUWx6PHnKdVzsRV-HJLKYybMuwRWYYV_tAyyELFHzMe4A1wigYxwDH4o1NCv2hnMn0_3h_O3rFIZ2YmL4eqjNH1IPgmkK0PFNW3tRl5Brn_ek2brOxJyw8OA3f-gzg9ou5qdF_UP0dMwLggsDNOgE23IZzu4NO4GuS-XPpiQGKQGBx4A8PW7UQaooDtAeInB1Eg43V5yhmHZdtUbULiEl6ZFFkknQ%26response_mode%3Dform_post%26nonce%3D639271137749547243.MTdlMTRmOWEtZjA2NS00NjE5LTk2ZjQtNjA5Y2M2ZTI2ZDJkNjg2NDY4MmEtYTFjNS00NDg1LTljN2UtMTNmNDk1ZTAxOTdh%26x-client-SKU%3DID_NET472%26x-client-ver%3D6.23.1.0"
 
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