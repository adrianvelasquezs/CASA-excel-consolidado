"""
Lectura de las plantillas por programa y escritura de la estructura común de
las hojas. Lo usan las dos versiones (onedrive y google_sheets).

Nada aquí asume rutas fijas: las plantillas llegan como una lista de archivos
o carpetas elegidos por el usuario.
"""

import re
from pathlib import Path

import openpyxl
from xlsxwriter.utility import xl_col_to_name

FILA_ENCABEZADO = 2  # índice 0 -> fila 3 de Excel, igual que en las plantillas
PRIMERA_FILA_DATOS = 3  # índice 0 -> fila 4 de Excel
FILAS_MIN, FILAS_MAX = 10, 2000

AZUL_RAW = "FF00B0F0"
AZUL = "#00B0F0"
VERDE = "#C6E0B4"
AZUL_CLARO = "#E6F7FD"
VERDE_CLARO = "#F0F7EB"
GRIS = "#F2F2F2"
GRIS_OSCURO = "#D9D9D9"

# Columnas que se llenan con fecha, listas cerradas o conteos.
COLUMNAS_FECHA = {"Fecha exacta inicio", "Fecha exacta de finalización"}
COLUMNAS_SI_NO = {
    "Montaje rúbricas Bloque Neón (SI / NO)": ["SI", "NO"],
    "Variabilidad en las notas? Responder Sí o No": ["Sí", "No"],
}
COLUMNAS_CONTEO = {
    "Número de competencias evaluadas",
    "Número de estudiantes inscritos curso",
    "Número de actividades evaluadas",
    "Número de estudiantes evaluados",
    "Número de criterios evaluados",
    "Número de objetivos evaluados",
}
# Columnas que identifican un curso: si alguna tiene datos, la fila cuenta como
# registrada (columna Programa y hoja Consolidado).
COLUMNAS_CLAVE = ("Código de curso", "CRN", "Curso")

HOJAS_RESERVADAS = {"instrucciones", "consolidado", "usuarios", "_auditoria"}


class ErrorGeneracion(Exception):
    """Error con un mensaje pensado para mostrarse tal cual al usuario."""


def carpeta_plantillas_incluidas():
    """Plantillas que vienen dentro de la app (o raw/ del repositorio al desarrollar). None si no hay."""
    empaquetadas = Path(__file__).resolve().parent.parent / "plantillas_incluidas"
    if empaquetadas.is_dir():
        return empaquetadas
    repositorio = Path(__file__).resolve().parents[3] / "raw"
    return repositorio if repositorio.is_dir() else None


def normalizar(texto):
    return re.sub(r"\s+", " ", str(texto)).strip()


def programa_de(archivo):
    """«AFIN Cronograma 2027 Bloque Neón.xlsx» -> «AFIN»."""
    return Path(archivo).stem.split(" ")[0].strip()


def expandir_rutas(rutas):
    """Convierte archivos y carpetas en la lista ordenada de plantillas .xlsx."""
    archivos = []
    for ruta in rutas:
        ruta = Path(ruta)
        if ruta.is_dir():
            archivos += sorted(p for p in ruta.glob("*.xlsx") if not p.name.startswith("~$"))
        elif ruta.is_file() and ruta.suffix.lower() == ".xlsx" and not ruta.name.startswith("~$"):
            archivos.append(ruta)
        elif not ruta.exists():
            raise ErrorGeneracion(f"No existe: {ruta}")
    if not archivos:
        raise ErrorGeneracion("No se eligió ninguna plantilla .xlsx.")
    return archivos


def validar_programa(programa, archivo):
    if not programa:
        raise ErrorGeneracion(f"No se pudo obtener el programa del nombre de archivo «{archivo.name}».")
    if len(programa) > 31 or re.search(r"[\[\]:*?/\\']", programa):
        raise ErrorGeneracion(f"«{programa}» (de {archivo.name}) no sirve como nombre de hoja: máximo 31 "
                              "caracteres y sin [ ] : * ? / \\ '")
    if programa.lower() in HOJAS_RESERVADAS:
        raise ErrorGeneracion(f"«{programa}» es un nombre de hoja reservado; renombre {archivo.name}.")


def leer_plantillas(rutas):
    """Devuelve ({programa: [(encabezado, es_azul), ...]}, {encabezado: ancho})."""
    plantillas = {}
    origen = {}
    anchos = {}
    for archivo in expandir_rutas(rutas):
        programa = programa_de(archivo)
        validar_programa(programa, archivo)
        if programa in plantillas:
            raise ErrorGeneracion(f"Hay dos plantillas del programa «{programa}»: {origen[programa].name} y {archivo.name}.")
        try:
            ws = openpyxl.load_workbook(archivo, read_only=False).worksheets[0]
        except Exception as error:  # archivo dañado, protegido o que no es Excel
            raise ErrorGeneracion(f"No se pudo leer {archivo.name}: {error}") from error
        columnas = []
        for celda in ws[FILA_ENCABEZADO + 1]:
            if celda.value is None:
                continue
            color = celda.fill.fgColor
            es_azul = bool(celda.fill.fill_type) and color.type == "rgb" and color.rgb == AZUL_RAW
            columnas.append((normalizar(celda.value), es_azul))
            ancho = ws.column_dimensions[celda.column_letter].width
            if ancho and normalizar(celda.value) not in anchos:
                anchos[normalizar(celda.value)] = ancho
        if "Programa" not in [n for n, _ in columnas]:
            raise ErrorGeneracion(f"{archivo.name} no tiene la fila de encabezados esperada (fila 3, columna «Programa»).")
        plantillas[programa] = columnas
        origen[programa] = archivo
    return plantillas, anchos


def unificar_columnas(plantillas):
    """Une los encabezados de todas las plantillas conservando el orden original. Devuelve (columnas, avisos)."""
    unificadas = []
    for columnas in plantillas.values():
        for i, (nombre, es_azul) in enumerate(columnas):
            existentes = [c[0] for c in unificadas]
            if nombre in existentes:
                previo = next(c for c in unificadas if c[0] == nombre)
                if previo[1] != es_azul:
                    raise ErrorGeneracion(f"La columna «{nombre}» tiene colores distintos entre plantillas.")
                continue
            # Insertar después de la columna que la precede en esta plantilla.
            anterior = columnas[i - 1][0] if i > 0 else None
            pos = existentes.index(anterior) + 1 if anterior in existentes else len(unificadas)
            unificadas.insert(pos, (nombre, es_azul))
    avisos = []
    for programa, cols in plantillas.items():
        faltan = [c for c, _ in unificadas if c not in {n for n, _ in cols}]
        if faltan:
            avisos.append(f"Aviso: {programa} no tenía {faltan}; se agregan para unificar.")
    return unificadas, avisos


def validar_filas(filas):
    if not isinstance(filas, int) or not FILAS_MIN <= filas <= FILAS_MAX:
        raise ErrorGeneracion(f"Las filas por programa deben estar entre {FILAS_MIN} y {FILAS_MAX}.")


def referencia_hoja(hoja):
    return "'" + hoja.replace("'", "''") + "'"


def crear_formatos(wb):
    base = {"border": 1, "border_color": "#BFBFBF", "valign": "top", "text_wrap": True}
    encabezado = {**base, "bold": True, "align": "center", "valign": "vcenter"}
    return {
        "titulo": wb.add_format({"bold": True, "font_size": 14}),
        "nota": wb.add_format({"italic": True, "font_color": "#595959"}),
        "leyenda_azul": wb.add_format({"bg_color": AZUL, "border": 1}),
        "leyenda_verde": wb.add_format({"bg_color": VERDE, "border": 1}),
        "leyenda_gris": wb.add_format({"bg_color": GRIS_OSCURO, "border": 1}),
        "enc_azul": wb.add_format({**encabezado, "bg_color": AZUL}),
        "enc_verde": wb.add_format({**encabezado, "bg_color": VERDE}),
        "enc_gris": wb.add_format({**encabezado, "bg_color": GRIS_OSCURO}),
        "dato_azul": wb.add_format({**base, "bg_color": AZUL_CLARO, "locked": False}),
        "dato_verde": wb.add_format({**base, "bg_color": VERDE_CLARO, "locked": False}),
        "fecha_azul": wb.add_format({**base, "bg_color": AZUL_CLARO, "locked": False, "num_format": "dd/mm/yyyy"}),
        "fecha_verde": wb.add_format({**base, "bg_color": VERDE_CLARO, "locked": False, "num_format": "dd/mm/yyyy"}),
        "bloqueado": wb.add_format({**base, "bg_color": GRIS, "locked": True}),
        "bloqueado_fecha": wb.add_format({**base, "bg_color": GRIS, "locked": True, "num_format": "dd/mm/yyyy hh:mm"}),
        "profesor_fijo": wb.add_format({**base, "bg_color": GRIS, "locked": True, "bold": True}),
        "fecha": wb.add_format({"num_format": "dd/mm/yyyy"}),
        "h1": wb.add_format({"bold": True, "font_size": 16}),
        "h2": wb.add_format({"bold": True, "font_size": 12, "top": 1}),
        "texto": wb.add_format({"text_wrap": True, "valign": "top"}),
        "enc_tabla": wb.add_format({"bold": True, "bg_color": GRIS_OSCURO, "border": 1}),
    }


def escribir_base_programa(wb, f, programa, columnas, anchos, filas, nota, columnas_sistema=()):
    """
    Crea la hoja de un programa con leyenda, encabezados, formato de celdas,
    columna Programa (fórmula) y validaciones. columnas_sistema se agregan al
    final en gris (las llena un script, no las personas).
    Devuelve (hoja, {nombre_columna: índice}).
    """
    ws = wb.add_worksheet(programa)
    ultima = PRIMERA_FILA_DATOS + filas - 1
    ws.write(0, 0, "Azul", f["leyenda_azul"])
    ws.write(0, 1, f"Programa / Profesores ({programa})")
    ws.write(1, 0, "Verde", f["leyenda_verde"])
    ws.write(1, 1, "Equipo CASA")
    ws.write(0, 4, f"Programa: {programa}", f["titulo"])
    ws.write(1, 4, nota, f["nota"])
    ws.set_row(0, 21.75)
    ws.set_row(1, 20.25)
    ws.set_row(FILA_ENCABEZADO, 45)

    idx = {nombre: i for i, (nombre, _) in enumerate(columnas)}
    col_programa = idx["Programa"]
    letras_clave = [xl_col_to_name(idx[c]) for c in COLUMNAS_CLAVE if c in idx]

    for c, (nombre, es_azul) in enumerate(columnas):
        ws.write(FILA_ENCABEZADO, c, nombre, f["enc_azul" if es_azul else "enc_verde"])
        ws.set_column(c, c, max(anchos.get(nombre, 20), 14))
        if c == col_programa:
            continue
        if nombre in COLUMNAS_FECHA:
            formato = f["fecha_azul" if es_azul else "fecha_verde"]
        else:
            formato = f["dato_azul" if es_azul else "dato_verde"]
        for r in range(PRIMERA_FILA_DATOS, ultima + 1):
            ws.write_blank(r, c, None, formato)

    for j, nombre in enumerate(columnas_sistema):
        c = len(columnas) + j
        idx[nombre] = c
        ws.write(FILA_ENCABEZADO, c, nombre, f["enc_gris"])
        ws.set_column(c, c, 30)
        formato = f["bloqueado_fecha"] if nombre.lower().startswith("fecha") else f["bloqueado"]
        for r in range(PRIMERA_FILA_DATOS, ultima + 1):
            ws.write_blank(r, c, None, formato)

    # Columna Programa: fórmula que se llena sola al registrar un curso.
    for r in range(PRIMERA_FILA_DATOS, ultima + 1):
        cond = "&".join(f"{l}{r + 1}" for l in letras_clave)
        ws.write_formula(r, col_programa, f'=IF({cond}="","","{programa}")', f["bloqueado"], "")

    for nombre, c in idx.items():
        if nombre in COLUMNAS_FECHA:
            ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {
                "validate": "date", "criteria": ">=", "value": "2020-01-01",
                "input_title": "Fecha", "input_message": "Formato dd/mm/aaaa",
                "error_message": "Ingrese una fecha válida (dd/mm/aaaa).",
            })
        elif nombre in COLUMNAS_SI_NO:
            ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {
                "validate": "list", "source": COLUMNAS_SI_NO[nombre],
            })
        elif nombre in COLUMNAS_CONTEO:
            ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {
                "validate": "integer", "criteria": ">=", "value": 0,
                "error_message": "Ingrese un número entero mayor o igual a 0.",
            })

    ws.autofilter(FILA_ENCABEZADO, 0, ultima, len(columnas) + len(columnas_sistema) - 1)
    ws.freeze_panes(PRIMERA_FILA_DATOS, idx.get("Profesor", 4) + 1)
    return ws, idx


def escribir_encabezados_consolidado(wb, f, columnas, anchos, columnas_sistema=()):
    ws = wb.add_worksheet("Consolidado")
    ws.write(0, 0, "Consolidado de todos los programas (solo lectura)", f["titulo"])
    ws.set_row(FILA_ENCABEZADO, 45)
    for c, (nombre, es_azul) in enumerate(columnas):
        ws.write(FILA_ENCABEZADO, c, nombre, f["enc_azul" if es_azul else "enc_verde"])
        ancho = max(anchos.get(nombre, 20), 14)
        ws.set_column(c, c, ancho, f["fecha"] if nombre in COLUMNAS_FECHA else None)
    for j, nombre in enumerate(columnas_sistema):
        ws.write(FILA_ENCABEZADO, len(columnas) + j, nombre, f["enc_gris"])
        ws.set_column(len(columnas) + j, len(columnas) + j, 30)
    nombres = [n for n, _ in columnas]
    ws.freeze_panes(PRIMERA_FILA_DATOS, nombres.index("Profesor") + 1 if "Profesor" in nombres else 5)
    return ws


def escribir_lineas(ws, f, lineas):
    ws.set_column(0, 0, 110)
    for r, (estilo, texto) in enumerate(lineas):
        ws.write(r, 0, texto, f[estilo])


def reemplazar_archivo(origen, destino):
    """Mueve el archivo temporal al destino elegido, con un mensaje claro si está abierto o sin permisos."""
    destino = Path(destino)
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        Path(origen).replace(destino)
    except OSError:
        try:
            import shutil
            shutil.move(str(origen), str(destino))  # otro disco o unidad
        except OSError as error:
            raise ErrorGeneracion(f"No se pudo guardar {destino.name} en {destino.parent}. "
                                  f"¿Está abierto en otro programa o la carpeta es de solo lectura? ({error})") from error
