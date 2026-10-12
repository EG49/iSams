"""
generar_formulas.py  -  PASO 2 (no abre el navegador, no modifica iSAMS)

Lee la hoja de la carpeta (ej. 'TI') de lectura_{promocion}.xlsx y calcula, para
cada materia, las fórmulas del trimestre. Todo queda en formulas_{promocion}.xlsx
para REVISARLO antes de escribirlo en iSAMS:

  SUMAINST = FLOOR((SUM(({ACTCT1}*0.20),({DEBT1}*0.20),...)),0.01)   (los % deben sumar 100)
  CALINST  = FLOOR(({SUMAINST1}*0.70),0.01)
  EXAT     = no se toca (nota plana)
  CALEXAT  = FLOOR(({EXAT1}*0.30),0.01)
  PROMTRI  = FLOOR(({CALEXAT1}+{CALINST1}),0.01)
{NOMBRE} es un marcador: al escribir en iSAMS se reemplaza por [ID<id real>] leído de la pestaña
'Columns' (los ids del árbol de estructura NO sirven: son números de fila que iSAMS reasigna).
FLOOR(x,0.01) trunca a centésimas (misma convención que las fórmulas actuales).

Las filas se listan en el orden en que hay que escribirlas (cada una depende de la anterior).
Materias que no cumplen (sin insumos con %, % que no suman 100, falta alguna fila) NO generan
fórmulas y van a la hoja 'problemas' (p. ej. las cualitativas con filas EQUI...).

Uso:  python generar_formulas.py 3BCH
"""

import sys

import openpyxl

from datos import (
    PROMOCION_POR_DEFECTO,
    ARCHIVO_SALIDA,
    ARCHIVO_FORMULAS,
    CARPETA_A_LEER,
    PREFIJO_SUMA_INSUMOS,
    PREFIJO_CAL_INSUMOS,
    PREFIJO_EXAMEN,
    PREFIJO_CAL_EXAMEN,
    PREFIJO_PROMEDIO_TRIMESTRE,
    PREFIJOS_NO_INSUMO,
    PORC_INSUMOS_DEFECTO,
    PORC_EXAMEN_DEFECTO,
    FLOOR_SIGNIFICANCIA,
)


def _factor(porcentaje):
    """20 -> '0.20', 15 -> '0.15', 70 -> '0.70', 12.5 -> '0.125' (mínimo 2 decimales)."""
    texto = f"{float(porcentaje) / 100:.6f}".rstrip("0")
    entero, _, decimales = texto.partition(".")
    return f"{entero}.{decimales.ljust(2, '0')}"


def _floor(expresion):
    """Trunca a centésimas como las fórmulas actuales: FLOOR((expr),0.01)."""
    return f"FLOOR(({expresion}),{FLOOR_SIGNIFICANCIA})"


def _ref(titulo):
    """Marcador {NOMBRE}: al escribir en iSAMS se reemplaza por [ID<id real de la columna>]."""
    return "{" + titulo + "}"


def _buscar(filas, prefijo):
    """Primera fila cuyo título empiece por el prefijo (ignora mayúsculas)."""
    for f in filas:
        if f["titulo"].upper().startswith(prefijo.upper()):
            return f
    return None


def formulas_de_materia(filas):
    """filas: [{'fila_id','titulo','porcentaje'}] de la carpeta, en orden.
    Devuelve (lista_formulas, error). lista_formulas = [(fila_dict, formula, nota, filas_referenciadas)]
    en el orden de escritura; si hay error, lista_formulas = [] y error es texto."""
    suma = _buscar(filas, PREFIJO_SUMA_INSUMOS)
    cal_ins = _buscar(filas, PREFIJO_CAL_INSUMOS)
    examen = _buscar(filas, PREFIJO_EXAMEN)
    cal_exa = _buscar(filas, PREFIJO_CAL_EXAMEN)
    prom = _buscar(filas, PREFIJO_PROMEDIO_TRIMESTRE)

    faltan = [n for n, f in [(PREFIJO_SUMA_INSUMOS, suma), (PREFIJO_CAL_INSUMOS, cal_ins),
                             (PREFIJO_EXAMEN, examen), (PREFIJO_CAL_EXAMEN, cal_exa),
                             (PREFIJO_PROMEDIO_TRIMESTRE, prom)] if f is None]
    if faltan:
        return [], f"faltan filas que empiecen por: {', '.join(faltan)}"

    insumos = [f for f in filas
               if f["porcentaje"] is not None
               and not f["titulo"].upper().startswith(tuple(p.upper() for p in PREFIJOS_NO_INSUMO))
               and not f["titulo"].upper().startswith(PREFIJO_EXAMEN.upper())]
    if not insumos:
        return [], "no hay columnas de insumo con % (¿materia cualitativa?)"

    total = sum(f["porcentaje"] for f in insumos)
    if abs(total - 100) > 1e-9:
        detalle = ", ".join(f"{f['titulo']}={f['porcentaje']}" for f in insumos)
        return [], f"los % de los insumos suman {total}, no 100 ({detalle})"

    p_ins = cal_ins["porcentaje"] if cal_ins["porcentaje"] is not None else PORC_INSUMOS_DEFECTO
    p_exa = cal_exa["porcentaje"] if cal_exa["porcentaje"] is not None else PORC_EXAMEN_DEFECTO
    if abs(p_ins + p_exa - 100) > 1e-9:
        return [], f"el % de insumos ({p_ins}) + el % del examen ({p_exa}) no suma 100"

    terminos = ",".join(f"({_ref(f['titulo'])}*{_factor(f['porcentaje'])})" for f in insumos)
    return [
        (suma, _floor(f"SUM({terminos})"), f"{len(insumos)} insumos, {total:g}%", insumos),
        (cal_ins, _floor(f"{_ref(suma['titulo'])}*{_factor(p_ins)}"), f"{suma['titulo']} x {p_ins:g}%", [suma]),
        (cal_exa, _floor(f"{_ref(examen['titulo'])}*{_factor(p_exa)}"), f"{examen['titulo']} x {p_exa:g}%", [examen]),
        (prom, _floor(f"{_ref(cal_exa['titulo'])}+{_ref(cal_ins['titulo'])}"), "examen + insumos, sobre 10", [cal_exa, cal_ins]),
    ], None


def main():
    promocion = (sys.argv[1] if len(sys.argv) > 1 else PROMOCION_POR_DEFECTO).upper()
    entrada = ARCHIVO_SALIDA.format(promocion=promocion)
    salida = ARCHIVO_FORMULAS.format(promocion=promocion)

    wb = openpyxl.load_workbook(entrada, data_only=True)
    hoja = wb[CARPETA_A_LEER]

    # Agrupar por materia conservando el orden: clave (grado, materia, id_gradebook)
    grupos = {}
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if not fila or fila[3] is None:
            continue
        grado, materia, id_gb, fila_id, titulo, porc = (list(fila) + [None] * 6)[:6]
        porc = float(porc) if porc not in (None, "") else None
        grupos.setdefault((grado, materia, str(id_gb)), []).append(
            {"fila_id": str(fila_id), "titulo": str(titulo or "").strip(), "porcentaje": porc})

    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "formulas"
    ws.append(["grado", "materia", "id_gradebook", "carpeta", "orden", "fila_destino", "formula", "nota", "referencias"])
    wp = out.create_sheet("problemas")
    wp.append(["grado", "materia", "id_gradebook", "problema"])

    ok = mal = 0
    for (grado, materia, id_gb), filas in grupos.items():
        formulas, error = formulas_de_materia(filas)
        if error:
            mal += 1
            wp.append([grado, materia, id_gb, error])
            continue
        ok += 1
        for orden, (f, formula, nota, refs) in enumerate(formulas, 1):
            referencias = ";".join(r["titulo"] for r in refs)
            ws.append([grado, materia, id_gb, CARPETA_A_LEER, orden, f["titulo"], formula, nota, referencias])

    ws.freeze_panes = "A2"
    wp.freeze_panes = "A2"
    out.save(salida)
    print(f"{promocion}: {ok} materia(s) con fórmulas, {mal} con problemas -> '{salida}'")
    if mal:
        print("Revisa la hoja 'problemas'.")


if __name__ == "__main__":
    main()