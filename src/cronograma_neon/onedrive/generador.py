"""
Versión Excel (OneDrive / SharePoint): libro consolidado protegido con contraseñas.

Modelo de permisos (protección de hoja + rangos con contraseña):
  - Encabezados, leyenda, columna «Programa», Instrucciones y Consolidado:
    bloqueados (solo la contraseña de administrador desbloquea).
  - Columnas azules de cada hoja: contraseña propia del programa.
  - Filas reservadas a un profesor: contraseña propia de ese profesor (el
    programa ya no puede editarlas).
  - Columnas verdes en todas las hojas: contraseña del equipo CASA.
  - Estructura del libro bloqueada (no se pueden borrar ni renombrar hojas).

Todas las rutas (plantillas, destino, profesores, contraseñas) las decide quien llama.
"""

import csv
import re
import secrets
import string
import tempfile
import zipfile
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
from xlsxwriter.worksheet import Worksheet

from ..comun.plantillas import (
    COLUMNAS_CLAVE,
    PRIMERA_FILA_DATOS,
    ErrorGeneracion,
    crear_formatos,
    escribir_base_programa,
    escribir_encabezados_consolidado,
    escribir_lineas,
    leer_plantillas,
    reemplazar_archivo,
    referencia_hoja,
    unificar_columnas,
    validar_filas,
)

NOMBRE_SUGERIDO = "Cronograma 2027 Bloque Neón - Consolidado.xlsx"
NOMBRE_SUGERIDO_CONTRASENAS = "contrasenas.csv"
ROLES = {"admin": "Administrador", "casa": "Equipo CASA", "programa": "Programa", "profesor": "Profesor"}
CLAVE_ADMIN = ("admin", "Libro completo")
CLAVE_CASA = ("casa", "Todas las hojas")

OPCIONES_PROTECCION = {
    "select_locked_cells": True,
    "select_unlocked_cells": True,
    "autofilter": True,
    "format_columns": True,
    "format_rows": True,
}

# ---------------------------------------------------------------------------
# Profesores y contraseñas
# ---------------------------------------------------------------------------


def leer_profesores_csv(ruta):
    """Lee un CSV programa,profesor,filas -> [(programa, profesor, filas)]."""
    try:
        with open(ruta, newline="", encoding="utf-8-sig") as f:
            lector = csv.DictReader(f)
            if not lector.fieldnames or not {"programa", "profesor", "filas"} <= set(lector.fieldnames):
                raise ErrorGeneracion(f"{Path(ruta).name} debe tener las columnas: programa, profesor, filas.")
            profesores = []
            for n, fila in enumerate(lector, start=2):
                programa, profesor = (fila["programa"] or "").strip(), (fila["profesor"] or "").strip()
                cantidad = (fila["filas"] or "").strip()
                if not programa and not profesor:
                    continue
                if not profesor or not cantidad.isdigit() or int(cantidad) < 1:
                    raise ErrorGeneracion(f"{Path(ruta).name} línea {n}: se requiere profesor y filas (entero > 0).")
                profesores.append((programa, profesor, int(cantidad)))
            return profesores
    except OSError as error:
        raise ErrorGeneracion(f"No se pudo leer {ruta}: {error}") from error


def _escribir_csv(ruta, encabezado, filas):
    try:
        Path(ruta).parent.mkdir(parents=True, exist_ok=True)
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f)
            escritor.writerow(encabezado)
            escritor.writerows(filas)
    except OSError as error:
        raise ErrorGeneracion(f"No se pudo guardar {ruta}: {error}") from error


def escribir_profesores_csv(ruta, profesores):
    _escribir_csv(ruta, ["programa", "profesor", "filas"], [(p, n, int(c)) for p, n, c in profesores])


def asignar_filas(programas, filas, profesores):
    """[(programa, profesor, filas)] -> {programa: [(profesor, fila_ini, fila_fin)]} (índices 0)."""
    asignaciones = {p: [] for p in programas}
    siguiente = {p: PRIMERA_FILA_DATOS for p in programas}
    for programa, profesor, cantidad in profesores:
        if programa not in asignaciones:
            raise ErrorGeneracion(f"El profesor «{profesor}» está asignado a «{programa}», que no está entre "
                                  f"las plantillas elegidas ({', '.join(programas)}).")
        if any(n.lower() == profesor.lower() for n, _, _ in asignaciones[programa]):
            raise ErrorGeneracion(f"«{profesor}» aparece dos veces en {programa}.")
        ini = siguiente[programa]
        fin = ini + int(cantidad) - 1
        if fin >= PRIMERA_FILA_DATOS + filas:
            raise ErrorGeneracion(f"Las filas reservadas en {programa} superan las {filas} filas por programa.")
        asignaciones[programa].append((profesor, ini, fin))
        siguiente[programa] = fin + 1
    return asignaciones


def generar_contrasena():
    alfabeto = "".join(c for c in string.ascii_letters + string.digits if c not in "0O1lI")
    return "".join(secrets.choice(alfabeto) for _ in range(10))


def leer_contrasenas_csv(ruta):
    """{(rol, ambito): contrasena} desde un CSV rol,ambito,contrasena."""
    try:
        with open(ruta, newline="", encoding="utf-8-sig") as f:
            lector = csv.DictReader(f)
            if not lector.fieldnames or not {"rol", "ambito", "contrasena"} <= set(lector.fieldnames):
                raise ErrorGeneracion(f"{Path(ruta).name} no es un archivo de contraseñas (columnas rol, ambito, contrasena).")
            return {(f["rol"], f["ambito"]): f["contrasena"] for f in lector if f["rol"] and f["contrasena"]}
    except OSError as error:
        raise ErrorGeneracion(f"No se pudo leer {ruta}: {error}") from error


def escribir_contrasenas_csv(ruta, contrasenas):
    _escribir_csv(ruta, ["rol", "ambito", "contrasena"], [[r, a, c] for (r, a), c in contrasenas.items()])


def claves_necesarias(programas, asignaciones):
    claves = [CLAVE_ADMIN, CLAVE_CASA] + [("programa", p) for p in programas]
    claves += [("profesor", f"{p} | {prof}") for p in programas for prof, _, _ in asignaciones[p]]
    return claves


def completar_contrasenas(existentes, claves):
    """Conserva las contraseñas existentes y crea las que falten. Devuelve (contrasenas, nuevas)."""
    resultado, nuevas = {}, 0
    for clave in claves:
        if existentes.get(clave):
            resultado[clave] = existentes[clave]
        else:
            resultado[clave] = generar_contrasena()
            nuevas += 1
    return resultado, nuevas


# ---------------------------------------------------------------------------
# Escritura del libro
# ---------------------------------------------------------------------------


def _rango_columnas(indices, fila_ini, fila_fin):
    """Convierte índices de columna en un sqref con áreas contiguas: 'B4:D10 F4:K10'."""
    areas, grupo = [], []
    for i in sorted(indices):
        if grupo and i != grupo[-1] + 1:
            areas.append(grupo)
            grupo = []
        grupo.append(i)
    if grupo:
        areas.append(grupo)
    return " ".join(f"{xl_rowcol_to_cell(fila_ini, g[0])}:{xl_rowcol_to_cell(fila_fin, g[-1])}" for g in areas)


def _nombre_rango(texto):
    return re.sub(r"[^A-Za-z0-9_]", "_", texto)[:200]


def _escribir_hoja_programa(wb, f, programa, columnas, anchos, filas, asignaciones, claves):
    ws, idx = escribir_base_programa(
        wb, f, programa, columnas, anchos, filas,
        "Gris = bloqueado. Al editar una celda de color, Excel pedirá la contraseña de su rol.",
    )
    ultima = PRIMERA_FILA_DATOS + filas - 1
    col_programa = idx["Programa"]
    col_profesor = idx.get("Profesor")
    azules = [i for i, (_, es_azul) in enumerate(columnas) if es_azul and i != col_programa]
    verdes = [i for i, (_, es_azul) in enumerate(columnas) if not es_azul]

    # Filas reservadas a profesores: el nombre queda fijo y el rango tiene su propia contraseña.
    filas_asignadas = set()
    for profesor, ini, fin in asignaciones:
        filas_asignadas.update(range(ini, fin + 1))
        editables = [i for i in azules if i != col_profesor]
        if col_profesor is not None:
            for r in range(ini, fin + 1):
                ws.write(r, col_profesor, profesor, f["profesor_fijo"])
        ws.unprotect_range(_rango_columnas(editables, ini, fin), _nombre_rango(f"Profesor_{programa}_{profesor}"),
                           claves[("profesor", f"{programa} | {profesor}")])

    # Filas libres del programa: bloques contiguos no asignados.
    libres = [r for r in range(PRIMERA_FILA_DATOS, ultima + 1) if r not in filas_asignadas]
    bloques, actual = [], []
    for r in libres:
        if actual and r != actual[-1] + 1:
            bloques.append(actual)
            actual = []
        actual.append(r)
    if actual:
        bloques.append(actual)
    if bloques:
        sqref = " ".join(_rango_columnas(azules, b[0], b[-1]) for b in bloques)
        ws.unprotect_range(sqref, _nombre_rango(f"Programa_{programa}"), claves[("programa", programa)])

    ws.unprotect_range(_rango_columnas(verdes, PRIMERA_FILA_DATOS, ultima), _nombre_rango(f"CASA_{programa}"),
                       claves[CLAVE_CASA])
    ws.protect(claves[CLAVE_ADMIN], OPCIONES_PROTECCION)


def _escribir_consolidado(wb, f, programas, columnas, anchos, filas, clave_admin):
    ws = escribir_encabezados_consolidado(wb, f, columnas, anchos)
    ws.write(1, 0, "Se actualiza automáticamente con lo registrado en cada hoja de programa. "
                   "Para filtrar u ordenar, copie y pegue como valores en otro libro.", f["nota"])
    ultima = PRIMERA_FILA_DATOS + filas  # fila Excel (1-based) de la última fila de datos
    ultima_col = xl_col_to_name(len(columnas) - 1)
    idx = {nombre: i for i, (nombre, _) in enumerate(columnas)}

    datos, condiciones = [], []
    for p in programas:
        h = referencia_hoja(p)
        rango = f"{h}!A{PRIMERA_FILA_DATOS + 1}:{ultima_col}{ultima}"
        datos.append(f'IF({rango}="","",{rango})')
        clave = "&".join(
            f"{h}!{xl_col_to_name(idx[c])}{PRIMERA_FILA_DATOS + 1}:{xl_col_to_name(idx[c])}{ultima}"
            for c in COLUMNAS_CLAVE if c in idx
        )
        condiciones.append(f'({clave})<>""')
    formula = f'=FILTER(VSTACK({",".join(datos)}),VSTACK({",".join(condiciones)}),"Sin registros")'
    if len(formula) > 8000:
        raise ErrorGeneracion("Son demasiadas plantillas para la fórmula del Consolidado (límite de Excel).")
    ws.write_dynamic_array_formula(PRIMERA_FILA_DATOS, 0, PRIMERA_FILA_DATOS, 0, formula)
    ws.protect(clave_admin, {"select_locked_cells": True, "format_columns": True})


def _escribir_instrucciones(wb, f, programas, asignaciones, clave_admin):
    ws = wb.add_worksheet("Instrucciones")
    lineas = [
        ("h1", "Cronograma 2027 Bloque Neón — Consolidado"),
        ("texto", "Este libro reúne en un solo archivo las plantillas de todos los programas. "
                  "Cada programa tiene su propia hoja; la hoja «Consolidado» las une automáticamente."),
        ("h2", "¿Qué puede editar cada quien?"),
        ("texto", "• Celdas AZULES (columnas del programa): solo el programa dueño de la hoja, con la "
                  "contraseña de su programa. Si una fila tiene el nombre del profesor fijo (en gris), "
                  "solo ese profesor puede editarla, con su contraseña personal."),
        ("texto", "• Celdas VERDES (columnas CASA): solo el equipo CASA, con la contraseña de CASA."),
        ("texto", "• Celdas GRISES, títulos, leyendas, la columna «Programa» y la hoja «Consolidado»: "
                  "nadie las edita (solo administración)."),
        ("h2", "¿Cómo se usa?"),
        ("texto", "1. Abra su hoja (pestaña con el nombre de su programa)."),
        ("texto", "2. Escriba en una celda de su color. Excel pedirá una contraseña: ingrese la de su rol. "
                  "Queda desbloqueado mientras tenga el archivo abierto."),
        ("texto", "3. La columna «Programa» se llena sola cuando registra el código, CRN o nombre del curso."),
        ("texto", "4. Las fechas van en formato dd/mm/aaaa; las columnas SI/NO tienen lista desplegable."),
        ("h2", "Programas incluidos"),
        ("texto", ", ".join(programas)),
    ]
    con_profesores = {p: a for p, a in asignaciones.items() if a}
    if con_profesores:
        lineas.append(("h2", "Filas asignadas a profesores"))
        for p, lista in con_profesores.items():
            detalle = "; ".join(f"{prof}: filas {ini + 1}–{fin + 1}" for prof, ini, fin in lista)
            lineas.append(("texto", f"{p} → {detalle}"))
    escribir_lineas(ws, f, lineas)
    ws.protect(clave_admin, {"select_locked_cells": True})


def _proteger_estructura(origen, destino, contrasena):
    """XlsxWriter no protege la estructura del libro: se inserta <workbookProtection> en workbook.xml."""
    hash_clave = Worksheet()._encode_password(contrasena)
    etiqueta = f'<workbookProtection workbookPassword="{hash_clave}" lockStructure="1"/>'
    with zipfile.ZipFile(origen) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            contenido = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                xml = contenido.decode("utf-8")
                if "<bookViews>" not in xml:
                    raise ErrorGeneracion("No se pudo proteger la estructura del libro (workbook.xml inesperado).")
                contenido = xml.replace("<bookViews>", etiqueta + "<bookViews>", 1).encode("utf-8")
            zout.writestr(item, contenido)


def generar_excel(rutas_plantillas, destino, filas=150, profesores=(), contrasenas=None):
    """
    Genera el Excel en `destino`. `contrasenas` ({(rol, ambito): clave}) se reutilizan; las que falten se crean.
    Devuelve (contrasenas_usadas, mensajes). Lanza ErrorGeneracion con un mensaje para el usuario.
    """
    validar_filas(filas)
    plantillas, anchos = leer_plantillas(rutas_plantillas)
    programas = list(plantillas)
    columnas, mensajes = unificar_columnas(plantillas)
    asignaciones = asignar_filas(programas, filas, list(profesores))
    claves, nuevas = completar_contrasenas(contrasenas or {}, claves_necesarias(programas, asignaciones))

    with tempfile.TemporaryDirectory() as tmp:
        sin_estructura = Path(tmp) / "libro.xlsx"
        final = Path(tmp) / "final.xlsx"
        wb = xlsxwriter.Workbook(str(sin_estructura), {"use_future_functions": True})
        f = crear_formatos(wb)
        _escribir_instrucciones(wb, f, programas, asignaciones, claves[CLAVE_ADMIN])
        for p in programas:
            _escribir_hoja_programa(wb, f, p, columnas, anchos, filas, asignaciones[p], claves)
        _escribir_consolidado(wb, f, programas, columnas, anchos, filas, claves[CLAVE_ADMIN])
        wb.close()
        _proteger_estructura(sin_estructura, final, claves[CLAVE_ADMIN])
        reemplazar_archivo(final, destino)

    mensajes.append(f"Excel guardado en: {destino}")
    mensajes.append(f"{len(programas)} programas, {len(columnas)} columnas, {filas} filas por programa, "
                    f"{sum(len(a) for a in asignaciones.values())} profesores con filas reservadas.")
    mensajes.append(f"Contraseñas: {len(claves)} en total, {nuevas} nuevas." if nuevas
                    else f"Contraseñas: se reutilizaron las {len(claves)} existentes.")
    return claves, mensajes
