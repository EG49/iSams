"""
generar_formulas.py  -  PASO 2 (no abre el navegador, no modifica iSAMS)

Lee la hoja de la carpeta (ej. 'TI') de lectura_{promocion}.xlsx y calcula, para
cada materia, las fórmulas del trimestre. Todo queda en formulas_{promocion}.xlsx
para REVISARLO antes de escribirlo en iSAMS:

  SUMAINST = FLOOR((SUM(({ACTCT1}*0.20),({DEBT1}*0.20),...)),0.01)   (los % deben sumar 100)
  CALINST  = FLOOR(({SUMAINST1}*0.70),0.01)
  EXAT     = no se toca (nota plana)
  CALEXAT  = FLOOR(({EXAT1}*0.30),0.01)
  PROMTRI  = FLOOR(SUM({CALINST1},{CALEXAT1}),0.01)
  EQUPRO   = escala A+..E- sobre {PROMTRI1} (solo si existe la fila EQUPRO...)
Cualitativas (un solo insumo, INSCUALT): CALINST sobre el insumo, CALEXAT, PROMTRI y todas las filas EQU... en escala A+..E-.
{NOMBRE} es un marcador: al escribir en iSAMS se reemplaza por [ID<id real>] leído de la pestaña
'Columns' (los ids del árbol de estructura NO sirven: son números de fila que iSAMS reasigna).
FLOOR(x,0.01) trunca a centésimas (misma convención que las fórmulas actuales).

Las filas se listan en el orden en que hay que escribirlas (cada una depende de la anterior).
Materias que no cumplen (sin insumos con %, % que no suman 100, falta alguna fila) NO generan
fórmulas y van a la hoja 'problemas' (p. ej. las cualitativas con filas EQUI...).

Uso:  python generar_formulas.py 3BCH [--solo cuantitativa | --solo cualitativa]
"""

import sys

import openpyxl

from funciones_isams import clasificar_materia

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
    PREFIJO_EQUIV_PROMEDIO,
    PREFIJO_INSUMO_CUALITATIVO,
    EQUIVALENCIAS_CUALITATIVA,
    PREFIJOS_NO_INSUMO,
    ALIAS_PREFIJOS,
    PORC_INSUMOS_DEFECTO,
    PORC_EXAMEN_DEFECTO,
    FLOOR_SIGNIFICANCIA,
)


def _factor(porcentaje):
    """20 -> '0.20', 15 -> '0.15', 70 -> '0.70', 12.5 -> '0.125' (mínimo 2 decimales)."""
    texto = f"{float(porcentaje) / 100:.6f}".rstrip("0")
    entero, _, decimales = texto.partition(".")
    return f"{entero}.{decimales.ljust(2, '0')}"


def _factor_corto(porcentaje):
    """70 -> '0.7' (así está escrito el 70% en 4EGB-A; el 30% va como '0.30')."""
    return f"{float(porcentaje) / 100:.6f}".rstrip("0").rstrip(".")


def _floor(expresion):
    """Trunca a centésimas como las fórmulas actuales: FLOOR((expr),0.01)."""
    return f"FLOOR(({expresion}),{FLOOR_SIGNIFICANCIA})"


_ESCALA = [(10, "A+"), (9, "A-"), (8, "B+"), (7, "B-"), (7, "B-"), (6, "C+"), (5, "C-"),
           (4, "D+"), (3, "D-"), (2, "E+")]  # igual a la fórmula de 4EGB-A (la línea de 7 va dos veces)


def _escala(ref):
    """Equivalencia en letras: IF(ISBLANK(x,true,false)," ",IF(x>=10,"A+",...,"E-")...)"""
    interior = '"E-"'
    for minimo, letra in reversed(_ESCALA):
        interior = f'IF({ref}>={minimo},"{letra}",{interior})'
    return f'IF(ISBLANK({ref},true,false)," ",{interior})'


def _ref(titulo):
    """Marcador {NOMBRE}: al escribir en iSAMS se reemplaza por [ID<id real de la columna>]."""
    return "{" + titulo + "}"


def coincide(titulo, prefijo):
    """¿El título empieza por el prefijo o por alguno de sus alias (datos.ALIAS_PREFIJOS)? Ignora mayúsculas."""
    validos = tuple(p.upper() for p in [prefijo] + list(ALIAS_PREFIJOS.get(prefijo, [])))
    return titulo.upper().startswith(validos)


def _buscar(filas, prefijo):
    """Primera fila cuyo título empiece por el prefijo o uno de sus alias (ignora mayúsculas)."""
    for f in filas:
        if coincide(f["titulo"], prefijo):
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
               and not any(coincide(f["titulo"], p) for p in PREFIJOS_NO_INSUMO)
               and not coincide(f["titulo"], PREFIJO_EXAMEN)]
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
    formulas = [
        (suma, _floor(f"SUM({terminos})"), f"{len(insumos)} insumos, {total:g}%", insumos),
        (cal_ins, _floor(f"{_ref(suma['titulo'])}*{_factor_corto(p_ins)}"), f"{suma['titulo']} x {p_ins:g}%", [suma]),
        (cal_exa, _floor(f"{_ref(examen['titulo'])}*{_factor(p_exa)}"), f"{examen['titulo']} x {p_exa:g}%", [examen]),
        (prom, f"FLOOR(SUM({_ref(cal_ins['titulo'])},{_ref(cal_exa['titulo'])}),{FLOOR_SIGNIFICANCIA})",
         "insumos + examen, sobre 10", [cal_ins, cal_exa]),
    ]
    equi = _buscar(filas, PREFIJO_EQUIV_PROMEDIO)
    if equi is not None:
        formulas.append((equi, _escala(_ref(prom["titulo"])), "equivalencia en letras del promedio", [prom]))
    return formulas, None


def formulas_cualitativa(filas):
    """Materia de un solo insumo (INSCUALT). Igual que 4EGB-A (Educación Física):
      CALINST = FLOOR(({INSCUALT}*0.7),0.01)      CALEXAT = FLOOR(({EXAT}*0.30),0.01)
      PROMTRI = FLOOR(SUM({CALINST},{CALEXAT}),0.01)
      cada fila EQU... = escala A+..E- sobre la columna de la que toma el valor.
    Devuelve (lista_formulas, error) con el mismo formato que formulas_de_materia."""
    insumo = _buscar(filas, PREFIJO_INSUMO_CUALITATIVO)
    cal_ins = _buscar(filas, PREFIJO_CAL_INSUMOS)
    examen = _buscar(filas, PREFIJO_EXAMEN)
    cal_exa = _buscar(filas, PREFIJO_CAL_EXAMEN)
    prom = _buscar(filas, PREFIJO_PROMEDIO_TRIMESTRE)
    faltan = [n for n, f in [(PREFIJO_INSUMO_CUALITATIVO, insumo), (PREFIJO_CAL_INSUMOS, cal_ins),
                             (PREFIJO_EXAMEN, examen), (PREFIJO_CAL_EXAMEN, cal_exa),
                             (PREFIJO_PROMEDIO_TRIMESTRE, prom)] if f is None]
    if faltan:
        return [], f"faltan filas que empiecen por: {', '.join(faltan)}"

    p_ins = cal_ins["porcentaje"] if cal_ins["porcentaje"] is not None else PORC_INSUMOS_DEFECTO
    p_exa = cal_exa["porcentaje"] if cal_exa["porcentaje"] is not None else PORC_EXAMEN_DEFECTO
    if abs(p_ins + p_exa - 100) > 1e-9:
        return [], f"el % de insumos ({p_ins}) + el % del examen ({p_exa}) no suma 100"

    formulas = [
        (cal_ins, _floor(f"{_ref(insumo['titulo'])}*{_factor_corto(p_ins)}"), f"{insumo['titulo']} x {p_ins:g}%", [insumo]),
        (cal_exa, _floor(f"{_ref(examen['titulo'])}*{_factor(p_exa)}"), f"{examen['titulo']} x {p_exa:g}%", [examen]),
        (prom, f"FLOOR(SUM({_ref(cal_ins['titulo'])},{_ref(cal_exa['titulo'])}),{FLOOR_SIGNIFICANCIA})",
         "insumos + examen, sobre 10", [cal_ins, cal_exa]),
    ]
    for prefijo_equi, prefijo_fuente in EQUIVALENCIAS_CUALITATIVA:
        equi = _buscar(filas, prefijo_equi)
        fuente = _buscar(filas, prefijo_fuente)
        if equi is not None and fuente is not None:
            formulas.append((equi, _escala(_ref(fuente["titulo"])),
                             f"equivalencia en letras de {fuente['titulo']}", [fuente]))
    return formulas, None


# Orden fijo de las 12 columnas de cada trimestre en las cualitativas "M-H" (5EGB-7EGB):
#  0 Insumo M-H (nota)            1 Equivalencia de la 0     2 Sumatoria Insumo = FLOOR({0},0.01)
#  3 Equivalencia de la 2         4 Calculo MH 70% = {2}*0.7  5 Equivalencia de la 4
#  6 Examen M-H (nota)            7 Equivalencia de la 6     8 Calculo EX = {6}*0.30
#  9 Equivalencia de la 8        10 PROMEDIO TRIMESTRE = {4}+{8}   11 Equivalencia de la 10
_MH_EQUIVALENCIAS = {1: 0, 3: 2, 5: 4, 7: 6, 9: 8, 11: 10}  # fila equivalente -> fila de la que toma el valor


# Variante de 10 columnas por trimestre (sin la fila "Sumatoria", como Educación Física 4EGB-A):
#  0 Insumo (nota)   1 Equivalencia de la 0   2 Calculo Insumo = {0}*0.7    3 Equivalencia de la 2
#  4 Examen (nota)   5 Equivalencia de la 4   6 Calculo Examen = {4}*0.30   7 Equivalencia de la 6
#  8 PROMEDIO TRIMESTRE = {2}+{6}             9 Equivalencia de la 8
_DIEZ_EQUIVALENCIAS = {1: 0, 3: 2, 5: 4, 7: 6, 9: 8}


def _formulas_cualitativa_10(filas):
    titulos = [f["titulo"] for f in filas]
    repetidos = sorted({t for t in titulos if titulos.count(t) > 1})
    if repetidos:
        return [], "nombres repetidos, renombra uno en iSAMS (no se sabría cuál escribir): " + ", ".join(repetidos)
    if not titulos[0].upper().startswith("INS") or not titulos[4].upper().startswith("EXA"):
        return [], "la posición 1 no es 'INS...' o la 5 no es 'EXA...': no coincide con el modelo de 10 columnas"
    p_ins, p_exa = PORC_INSUMOS_DEFECTO, PORC_EXAMEN_DEFECTO
    r = lambda i: _ref(titulos[i])
    formulas = [
        (filas[2], _floor(f"{r(0)}*{_factor_corto(p_ins)}"), f"{titulos[0]} x {p_ins:g}%", [filas[0]]),
        (filas[6], _floor(f"{r(4)}*{_factor(p_exa)}"), f"{titulos[4]} x {p_exa:g}%", [filas[4]]),
        (filas[8], f"FLOOR(SUM({r(2)},{r(6)}),{FLOOR_SIGNIFICANCIA})", "insumos + examen, sobre 10", [filas[2], filas[6]]),
    ]
    for destino, fuente in _DIEZ_EQUIVALENCIAS.items():
        formulas.append((filas[destino], _escala(r(fuente)), f"equivalencia en letras de {titulos[fuente]}", [filas[fuente]]))
    return formulas, None


def formulas_cualitativa_mh(filas):
    """Cualitativas de 5EGB-7EGB. Las columnas se identifican por su POSICIÓN en la carpeta (sus nombres
    no son estándar y a veces se repiten): 12 columnas (modelo M-H) o 10 (sin la fila Sumatoria).
    Mismo formato que formulas_de_materia."""
    if len(filas) == 10:
        return _formulas_cualitativa_10(filas)
    if len(filas) != 12:
        return [], f"se esperaban 12 (o 10) columnas en la carpeta y hay {len(filas)}"
    titulos = [f["titulo"] for f in filas]
    repetidos = sorted({t for t in titulos if titulos.count(t) > 1})
    if repetidos:
        pos = ", ".join(f"{t} (posiciones {[i + 1 for i, x in enumerate(titulos) if x == t]})" for t in repetidos)
        return [], f"nombres repetidos, renombra uno en iSAMS (no se sabría cuál escribir): {pos}"
    if not titulos[0].upper().startswith("INSUMO") or not titulos[6].upper().startswith("EXAMEN"):
        return [], "la posición 1 no es 'Insumo...' o la 7 no es 'Examen...': no coincide con el modelo M-H"
    p_ins, p_exa = PORC_INSUMOS_DEFECTO, PORC_EXAMEN_DEFECTO
    r = lambda i: _ref(titulos[i])
    formulas = [
        (filas[2], f"FLOOR({r(0)},{FLOOR_SIGNIFICANCIA})", f"{titulos[0]} (un solo insumo)", [filas[0]]),
        (filas[4], _floor(f"{r(2)}*{_factor_corto(p_ins)}"), f"{titulos[2]} x {p_ins:g}%", [filas[2]]),
        (filas[8], _floor(f"{r(6)}*{_factor(p_exa)}"), f"{titulos[6]} x {p_exa:g}%", [filas[6]]),
        (filas[10], f"FLOOR(SUM({r(4)},{r(8)}),{FLOOR_SIGNIFICANCIA})", "insumos + examen, sobre 10", [filas[4], filas[8]]),
    ]
    for destino, fuente in _MH_EQUIVALENCIAS.items():
        formulas.append((filas[destino], _escala(r(fuente)), f"equivalencia en letras de {titulos[fuente]}", [filas[fuente]]))
    return formulas, None


def main():
    args = sys.argv[1:]
    solo = args[args.index("--solo") + 1].lower() if "--solo" in args else None
    if solo not in (None, "cuantitativa", "cualitativa"):
        sys.exit("--solo admite: cuantitativa | cualitativa (incluye las cualitativas M-H)")
    pos = [a for i, a in enumerate(args) if not a.startswith("--") and not (i > 0 and args[i - 1] == "--solo")]
    promocion = (pos[0] if pos else PROMOCION_POR_DEFECTO).upper()
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
    wo = out.create_sheet("omitidas")
    wo.append(["grado", "materia", "id_gradebook", "categoria", "nota"])
    notas = {
        "hija": "hija de Educación Cultural y Artística: proceso aparte (pendiente)",
        "madre": "promedio de sus hijas (Art y Música): proceso aparte (pendiente)",
        "manual": "se llena a mano, no se toca",
        "excluida": "no es materia de este proceso",
    }

    ok = mal = 0
    omitidas = {}
    for (grado, materia, id_gb), filas in grupos.items():
        categoria = clasificar_materia(grado, materia)
        if solo and categoria in ("cuantitativa", "cualitativa", "cualitativa_mh") \
                and (categoria == "cuantitativa") != (solo == "cuantitativa"):
            omitidas["fuera de --solo"] = omitidas.get("fuera de --solo", 0) + 1
            wo.append([grado, materia, id_gb, categoria, f"fuera del filtro --solo {solo}"])
            continue
        if categoria not in ("cuantitativa", "cualitativa", "cualitativa_mh"):
            omitidas[categoria] = omitidas.get(categoria, 0) + 1
            wo.append([grado, materia, id_gb, categoria, notas.get(categoria, "")])
            continue
        if categoria == "cualitativa":
            formulas, error = formulas_cualitativa(filas)
        elif categoria == "cualitativa_mh":
            formulas, error = formulas_cualitativa_mh(filas)
        else:
            formulas, error = formulas_de_materia(filas)
        if error:
            mal += 1
            wp.append([grado, materia, id_gb, f"{error}  |  columnas que hay en {CARPETA_A_LEER}: "
                       + (", ".join(f["titulo"] for f in filas) or "(ninguna)")])
            continue
        ok += 1
        for orden, (f, formula, nota, refs) in enumerate(formulas, 1):
            referencias = ";".join(r["titulo"] for r in refs)
            ws.append([grado, materia, id_gb, CARPETA_A_LEER, orden, f["titulo"], formula, nota, referencias])

    ws.freeze_panes = "A2"
    wp.freeze_panes = "A2"
    wo.freeze_panes = "A2"
    out.save(salida)
    print(f"{promocion}: {ok} materia(s) con fórmulas, {mal} con problemas -> '{salida}'")
    if omitidas:
        print("Omitidas (hoja 'omitidas'): " + ", ".join(f"{n} {c}" for c, n in omitidas.items()))
    if mal:
        print("Revisa la hoja 'problemas'.")


if __name__ == "__main__":
    main()