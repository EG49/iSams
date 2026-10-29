"""
renombrar_duplicadas.py  -  corrige el nombre repetido de las cualitativas M-H (5EGB-7EGB)

En esas materias, en TRIMESTRE II, las posiciones 8 y 10 se llaman las dos "Equi Examen M-H T-II".
La de la posición 10 es la equivalencia del "Calculo EX", así que este script le pone:
    Name        = "Equivalencia Examen 30% M-H T-<romano>"
    Short Name  = "ECEXAMAHT<n>"            (n = número del trimestre)
La fórmula no se toca. Solo actúa si la carpeta tiene 12 columnas y las posiciones 8 y 10 tienen
el mismo nombre que empieza por "Equi Examen". Lo demás se deja como está.

Por defecto NO abre el navegador: solo lista lo que cambiaría (renombres_{promocion}.xlsx).
Con --ejecutar lo hace en iSAMS (login manual), una materia a la vez (un solo guardado por materia).

Uso:  python renombrar_duplicadas.py 5EGB [6EGB 7EGB ...] [--ejecutar] [--id 2326] [--max N]
"""

import sys
import time

import openpyxl

from datos import LOGIN_URL, RUTA_MODULO_GRADEBOOKS, CARPETA_A_LEER, EXCLUIR_MATERIAS, TRIMESTRES
from funciones_isams import (
    crear_driver, login_manual, cerrar_alerta_pendiente, leer_estructura, filas_de_carpeta,
    clasificar_materia, normalizar, analizar_fila, buscar_elementos_con_scroll, buscar_boton_editar_columna, clic_seguro,
    elemento_profundo, clic_next_wizard_verificado, guardar_wizard_calculo, poner_texto_form,
)


def planificar(hijas, nuevo_nombre, nuevo_corto):
    """hijas: [(fila_id, fila_dict)] de la carpeta. Devuelve (fila_id, titulo_actual, corto_actual, motivo)
    de la columna de la posición 10 si hay que corregirla, o None. Hay que corregirla cuando:
      - su nombre está repetido con el de la posición 8 ("Equi Examen ..."), o
      - ya tiene el Name nuevo pero su Short Name no es el nuevo (se cambió solo el Name)."""
    if len(hijas) != 12:
        return None
    titulos = [f["titulo"] for _fid, f in hijas]
    fid, fila = hijas[9]
    corto = analizar_fila(fila)[0]
    if titulos[7] == titulos[9] and normalizar(titulos[9]).startswith(normalizar("Equi Examen")):
        return fid, titulos[9], corto, "nombre repetido"
    if normalizar(titulos[9]) == normalizar(nuevo_nombre) and normalizar(corto) != normalizar(nuevo_corto):
        return fid, titulos[9], corto, "falta cambiar el Short Name"
    return None


def renombrar(driver, fila_id, nuevo_nombre, nuevo_corto):
    filas = buscar_elementos_con_scroll(driver, f"//div[contains(@class,'row')][@id='{fila_id}']", tiempo_espera=45)
    if not filas:
        raise TimeoutError(f"no se encontró la fila con id {fila_id}")
    clic_seguro(driver, buscar_boton_editar_columna(driver, filas[0]))
    time.sleep(1.5)
    form = elemento_profundo(driver, "isams-data-form#calculationPropertiesForm", 20)
    poner_texto_form(driver, form, "columnName", nuevo_nombre)
    poner_texto_form(driver, form, "columnShortName", nuevo_corto)
    time.sleep(0.5)
    clic_next_wizard_verificado(driver, "isams-wizard#calculationWizard")  # Step 2: la fórmula se deja igual
    guardar_wizard_calculo(driver)


def main():
    args = sys.argv[1:]
    ejecutar = "--ejecutar" in args
    maximo = int(args[args.index("--max") + 1]) if "--max" in args else None
    solo_id = args[args.index("--id") + 1] if "--id" in args else None
    pos = [a for i, a in enumerate(args) if not a.startswith("--") and not (i > 0 and args[i - 1] in ("--max", "--id"))]
    if not pos:
        print("Indica la promoción o curso, ej.: python renombrar_duplicadas.py 5EGB")
        return
    etiquetas = [p.upper() for p in pos]
    etiqueta = "_".join(etiquetas)
    clave = CARPETA_A_LEER.upper().replace(" ", "")
    if clave not in TRIMESTRES:
        sys.exit(f"CARPETA_A_LEER='{CARPETA_A_LEER}' no está en TRIMESTRES (datos.py).")
    n, romano = TRIMESTRES[clave]
    nuevo_nombre = f"Equivalencia Examen 30% M-H T-{romano}"
    nuevo_corto = f"ECEXAMAHT{n}"

    from leer_materias import cargar_materias
    from escribir_formulas import abrir_materia
    excluir = {normalizar(m) for m in EXCLUIR_MATERIAS}
    materias = [m for e in etiquetas for m in cargar_materias(e) if normalizar(m[2]) not in excluir
                and clasificar_materia(m[1], m[2]) == "cualitativa_mh"]
    if solo_id:
        materias = [m for m in materias if m[0] == solo_id]
    if maximo:
        materias = materias[:maximo]
    if not materias:
        print(f"No hay materias cualitativas M-H en '{etiqueta}'.")
        return
    print(f"{etiqueta} / {CARPETA_A_LEER}: {len(materias)} materia(s) M-H a revisar" + ("." if ejecutar else " (solo plan)."))

    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "renombres"
    ws.append(["grado", "materia", "id_gradebook", "carpeta", "nombre_actual", "nombre_nuevo", "short_nuevo", "estado"])
    ruta = f"renombres_{etiqueta}_{CARPETA_A_LEER.upper()}.xlsx"

    driver = crear_driver()
    login_manual(driver)  # hay que leer la estructura de cada materia, así que siempre se abre el navegador
    driver.get(LOGIN_URL.rstrip("/") + "/" + RUTA_MODULO_GRADEBOOKS)
    time.sleep(4)
    cambios = errores = 0
    try:
        for i, (id_gb, grado, materia) in enumerate(materias, 1):
            print(f"[{i}/{len(materias)}] {grado} {materia} (id={id_gb})")
            try:
                abrir_materia(driver, id_gb)
                hijas = filas_de_carpeta(leer_estructura(driver), CARPETA_A_LEER)
                if hijas is None:
                    raise RuntimeError(f"no tiene la carpeta '{CARPETA_A_LEER}'")
                objetivo = planificar(hijas, nuevo_nombre, nuevo_corto)
                if objetivo is None:
                    print("   nada que renombrar.")
                    ws.append([grado, materia, id_gb, CARPETA_A_LEER, "", "", "", "sin cambios"])
                    continue
                fila_id, actual, corto_actual, motivo = objetivo
                if not ejecutar:
                    print(f"   cambiaría la posición 10 ({motivo}): '{actual}' [{corto_actual}] -> '{nuevo_nombre}' [{nuevo_corto}]")
                    ws.append([grado, materia, id_gb, CARPETA_A_LEER, actual, nuevo_nombre, nuevo_corto, f"pendiente ({motivo})"])
                    continue
                renombrar(driver, fila_id, nuevo_nombre, nuevo_corto)
                # comprobación: releer y exigir que la posición 10 tenga el Name Y el Short Name nuevos
                abrir_materia(driver, id_gb)
                despues = filas_de_carpeta(leer_estructura(driver), CARPETA_A_LEER)
                if len(despues) != 12:
                    raise RuntimeError(f"tras guardar, la carpeta tiene {len(despues)} columnas (se esperaban 12)")
                titulos = [f["titulo"] for _fid, f in despues]
                corto = analizar_fila(despues[9][1])[0]
                if normalizar(titulos[9]) != normalizar(nuevo_nombre):
                    raise RuntimeError(f"el Name quedó '{titulos[9]}', no '{nuevo_nombre}'")
                if normalizar(corto) != normalizar(nuevo_corto):
                    raise RuntimeError(f"el Short Name quedó '{corto}', no '{nuevo_corto}'")
                if titulos.count(titulos[7]) != 1:
                    raise RuntimeError("siguen habiendo nombres repetidos")
                print("   ✅ Name y Short Name correctos.")
                cambios += 1
                ws.append([grado, materia, id_gb, CARPETA_A_LEER, actual, nuevo_nombre, nuevo_corto, "ok"])
            except Exception as e:
                errores += 1
                print(f"   ⚠️  ERROR - {e}")
                ws.append([grado, materia, id_gb, CARPETA_A_LEER, "", "", "", f"error: {e}"])
                cerrar_alerta_pendiente(driver)
            out.save(ruta)
    except KeyboardInterrupt:
        print("\nInterrumpido.")
    finally:
        out.save(ruta)
    print("\n" + "=" * 60)
    print(f"Listo. {cambios} renombrada(s), {errores} con error. Detalle en '{ruta}'.")
    input("\nPresiona ENTER para cerrar el navegador...")
    driver.quit()


if __name__ == "__main__":
    main()