"""
escribir_formulas.py  -  PASO 3 (ESCRIBE en iSAMS)

Lee formulas_{promocion}.xlsx (generado por generar_formulas.py y revisado por ti)
y, materia por materia, abre cada columna calculada y escribe su fórmula:

    ubicar la fila por su ID en el árbol -> lápiz (editar) -> Step 1: Formula Builder +
    Not Shared -> Next -> escribir la fórmula -> Next, Next, Next -> Finish

Seguridad:
  * VERIFICACIÓN INICIAL: antes de escribir la primera fórmula comprueba en la pestaña
    'Columns' que el ID de fila del árbol coincide con el ID que usa [ID...] en las
    fórmulas. Si no coincide, se detiene TODO (el error sería sistemático).
  * Regla de siempre: si una fórmula falla, se reporta y se continúa con la siguiente.
  * resultado_{promocion}.xlsx se guarda tras CADA fórmula (ok / error). Con --reanudar
    se saltan las que ya quedaron 'ok'.

Uso:
    python escribir_formulas.py 3BCH                 # todas las materias
    python escribir_formulas.py 3BCH --max 1         # solo las primeras N materias (para probar)
    python escribir_formulas.py 3BCH --reanudar      # continúa lo que quedó pendiente/con error
"""

import sys
import time

import openpyxl

from datos import (
    LOGIN_URL,
    RUTA_MODULO_GRADEBOOKS,
    RUTA_EDITAR_GRADEBOOK,
    PROMOCION_POR_DEFECTO,
    ARCHIVO_FORMULAS,
    ARCHIVO_RESULTADO,
)
from funciones_isams import (
    crear_driver,
    login_manual,
    cerrar_alerta_pendiente,
    esperar_filas,
    buscar_elementos_con_scroll,
    buscar_boton_editar_columna,
    clic_seguro,
    configurar_propiedades_calculo,
    abrir_pestana_columns,
    recolectar_ids_columnas,
    escribir_formula,
    guardar_wizard_calculo,
)


def cargar_formulas(promocion):
    """{(grado, materia, id_gb): [dict(orden, fila_id, titulo, formula, referencias)]}"""
    wb = openpyxl.load_workbook(ARCHIVO_FORMULAS.format(promocion=promocion), data_only=True)
    grupos = {}
    for fila in wb["formulas"].iter_rows(min_row=2, values_only=True):
        if not fila or fila[0] is None:
            continue
        grado, materia, id_gb, orden, fila_id, titulo, formula, _nota, refs = (list(fila) + [None] * 9)[:9]
        grupos.setdefault((grado, materia, str(id_gb)), []).append({
            "orden": int(orden), "fila_id": str(fila_id), "titulo": str(titulo),
            "formula": str(formula), "referencias": str(refs or ""),
        })
    for lista in grupos.values():
        lista.sort(key=lambda d: d["orden"])
    return grupos


def cargar_resultados_ok(ruta):
    """Set de (id_gb, fila_id) que ya quedaron 'ok' en una corrida anterior."""
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True)
    except Exception:
        return set(), []
    filas, ok = [], set()
    for f in wb.active.iter_rows(min_row=2, values_only=True):
        if not f or f[0] is None:
            continue
        filas.append(list(f))
        if f[6] == "ok":
            ok.add((str(f[2]), str(f[3])))
    return ok, filas


def guardar_resultados(ruta, filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "resultado"
    ws.append(["grado", "materia", "id_gradebook", "fila_id", "fila", "formula", "estado", "detalle"])
    for r in filas:
        ws.append(r)
    ws.freeze_panes = "A2"
    wb.save(ruta)


def abrir_materia(driver, id_gb):
    url = LOGIN_URL.rstrip("/") + "/" + RUTA_EDITAR_GRADEBOOK.format(id=id_gb)
    driver.get(url)  # mismo path del módulo: solo cambia el #hash (ruta de la SPA)
    time.sleep(2)
    cerrar_alerta_pendiente(driver)
    if not esperar_filas(driver, 25):
        driver.get(LOGIN_URL.rstrip("/") + "/" + RUTA_MODULO_GRADEBOOKS)
        time.sleep(3)
        driver.get(url)
        time.sleep(2)
        if not esperar_filas(driver, 25):
            raise TimeoutError(f"el árbol no cargó. URL actual: {driver.current_url}")


def ubicar_fila_por_id(driver, fila_id, tiempo_espera=45):
    filas = buscar_elementos_con_scroll(
        driver, f"//div[contains(@class,'row')][@id='{fila_id}']", tiempo_espera=tiempo_espera)
    if not filas:
        raise TimeoutError(f"no se encontró la fila con id={fila_id} en el árbol")
    return filas[0]


def verificar_ids(driver, referencias):
    """Con el wizard abierto en el Step 2: confirma que el ID de fila coincide con el ID de
    la pestaña 'Columns' (el que se usa en [ID...]). Toma la primera referencia 'id=titulo'."""
    primera = referencias.split(";")[0]
    fila_id, _, titulo = primera.partition("=")
    abrir_pestana_columns(driver)
    ids = recolectar_ids_columnas(driver, titulo)
    if fila_id not in ids:
        raise RuntimeError(
            f"VERIFICACIÓN FALLIDA: la columna '{titulo}' tiene id de fila {fila_id}, pero en la pestaña "
            f"'Columns' sus IDs son {ids}. El id del árbol NO sirve para las fórmulas; no se escribe nada.")


def escribir_una(driver, f, verificar):
    fila = ubicar_fila_por_id(driver, f["fila_id"])
    clic_seguro(driver, buscar_boton_editar_columna(driver, fila))
    time.sleep(1.5)
    configurar_propiedades_calculo(driver)
    if verificar:
        verificar_ids(driver, f["referencias"])
    escribir_formula(driver, f["formula"])
    guardar_wizard_calculo(driver)


def main():
    args = sys.argv[1:]
    reanudar = "--reanudar" in args
    maximo = None
    if "--max" in args:
        maximo = int(args[args.index("--max") + 1])
    posicionales = [a for i, a in enumerate(args)
                    if not a.startswith("--") and not (i > 0 and args[i - 1] == "--max")]
    promocion = (posicionales[0] if posicionales else PROMOCION_POR_DEFECTO).upper()

    grupos = cargar_formulas(promocion)
    if maximo:
        grupos = dict(list(grupos.items())[:maximo])
    total = sum(len(v) for v in grupos.values())
    print(f"{promocion}: {len(grupos)} materia(s), {total} fórmula(s) a escribir.")

    ruta_res = ARCHIVO_RESULTADO.format(promocion=promocion)
    ya_ok, resultados = cargar_resultados_ok(ruta_res) if reanudar else (set(), [])
    if reanudar:
        # quitar las filas con error previo: se reintentan y se vuelven a registrar
        resultados = [r for r in resultados if r[6] == "ok"]
        print(f"Reanudando: {len(ya_ok)} fórmula(s) ya estaban ok.")

    driver = crear_driver()
    login_manual(driver)
    driver.get(LOGIN_URL.rstrip("/") + "/" + RUTA_MODULO_GRADEBOOKS)
    time.sleep(4)

    verificado = False
    errores = []
    try:
        for i, ((grado, materia, id_gb), lista) in enumerate(grupos.items(), 1):
            print(f"\n[{i}/{len(grupos)}] {grado} {materia} (id={id_gb})")
            pendientes = [f for f in lista if (id_gb, f["fila_id"]) not in ya_ok]
            if not pendientes:
                print("   todo ok de corridas anteriores, se salta.")
                continue
            try:
                abrir_materia(driver, id_gb)
            except Exception as e:
                msg = f"{grado} {materia}: ERROR al abrir - {e}"
                print(f"⚠️  {msg}")
                errores.append(msg)
                for f in pendientes:
                    resultados.append([grado, materia, id_gb, f["fila_id"], f["titulo"], f["formula"], "error", f"no abrió: {e}"])
                guardar_resultados(ruta_res, resultados)
                cerrar_alerta_pendiente(driver)
                continue

            for f in pendientes:
                try:
                    escribir_una(driver, f, verificar=not verificado)
                    verificado = True
                    print(f"✅ {f['titulo']}: {f['formula']}")
                    resultados.append([grado, materia, id_gb, f["fila_id"], f["titulo"], f["formula"], "ok", ""])
                except RuntimeError as e:
                    if str(e).startswith("VERIFICACIÓN FALLIDA"):
                        print(f"\n⛔ {e}")
                        raise SystemExit(1)
                    raise
                except Exception as e:
                    msg = f"{grado} {materia} / {f['titulo']}: ERROR - {e}"
                    print(f"⚠️  {msg}")
                    errores.append(msg)
                    resultados.append([grado, materia, id_gb, f["fila_id"], f["titulo"], f["formula"], "error", str(e)])
                    cerrar_alerta_pendiente(driver)
                    # Si quedó el wizard abierto, recargamos la materia para dejarla limpia.
                    try:
                        abrir_materia(driver, id_gb)
                    except Exception:
                        pass
                guardar_resultados(ruta_res, resultados)
    except KeyboardInterrupt:
        print("\nInterrumpido; lo escrito hasta ahora está en", ruta_res)
    finally:
        guardar_resultados(ruta_res, resultados)

    ok = sum(1 for r in resultados if r[6] == "ok")
    print("\n" + "=" * 60)
    print(f"Listo. {ok} fórmula(s) ok, {len(errores)} con error. Detalle en '{ruta_res}'.")
    for e in errores:
        print(f"  - {e}")
    input("\nPresiona ENTER para cerrar el navegador...")
    driver.quit()


if __name__ == "__main__":
    main()