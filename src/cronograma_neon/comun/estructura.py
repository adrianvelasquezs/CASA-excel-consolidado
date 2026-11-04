"""
Escritura de la estructura del libro a partir de la Configuracion: hojas de
programa, encabezados del Consolidado y hoja «Permisos».
"""

import shutil
from pathlib import Path

from xlsxwriter.utility import xl_col_to_name

from .configuracion import (
    COLUMNA_PROGRAMA,
    COLUMNAS_SISTEMA_FIN,
    TIPOS,
    ErrorGeneracion,
    opciones_de,
)

FILA_ENCABEZADO = 2  # índice 0 -> fila 3 de la hoja
PRIMERA_FILA_DATOS = 3  # índice 0 -> fila 4 de la hoja

# Color del encabezado y de las celdas según quién edita la columna.
COLORES = {
    "programa": ("#00B0F0", "#E6F7FD", "Azul", "Programa (profesores / gestores)"),
    "casa": ("#C6E0B4", "#F0F7EB", "Verde", "Equipo CASA"),
    "mixto": ("#FFE699", "#FFF7DD", "Amarillo", "Programa y CASA"),
    "admin": ("#D9D9D9", "#F2F2F2", "Gris", "Solo administración"),
}


def crear_formatos(wb):
    base = {"border": 1, "border_color": "#BFBFBF", "valign": "top", "text_wrap": True}
    encabezado = {**base, "bold": True, "align": "center", "valign": "vcenter"}
    f = {
        "titulo": wb.add_format({"bold": True, "font_size": 14}),
        "nota": wb.add_format({"italic": True, "font_color": "#595959"}),
        "fecha_hora": wb.add_format({"num_format": "dd/mm/yyyy hh:mm"}),
        "fecha": wb.add_format({"num_format": "dd/mm/yyyy"}),
        "h1": wb.add_format({"bold": True, "font_size": 16}),
        "h2": wb.add_format({"bold": True, "font_size": 12, "top": 1}),
        "texto": wb.add_format({"text_wrap": True, "valign": "top"}),
        "enc_tabla": wb.add_format({"bold": True, "bg_color": COLORES["admin"][0], "border": 1, "text_wrap": True}),
        "celda_tabla": wb.add_format({"border": 1, "border_color": "#BFBFBF", "text_wrap": True, "valign": "top"}),
    }
    for clave, (fuerte, claro, _, _) in COLORES.items():
        f[f"leyenda_{clave}"] = wb.add_format({"bg_color": fuerte, "border": 1})
        f[f"enc_{clave}"] = wb.add_format({**encabezado, "bg_color": fuerte})
        f[f"dato_{clave}"] = wb.add_format({**base, "bg_color": claro})
        f[f"fecha_{clave}"] = wb.add_format({**base, "bg_color": claro, "num_format": "dd/mm/yyyy"})
    f["sistema"] = f["dato_admin"]
    f["sistema_fecha"] = wb.add_format({**base, "bg_color": COLORES["admin"][1], "num_format": "dd/mm/yyyy hh:mm"})
    return f


def indices(configuracion):
    """{nombre: índice} de todas las columnas de una hoja de programa, incluidas las del sistema."""
    nombres = [COLUMNA_PROGRAMA] + [c.nombre for c in configuracion.columnas] + list(COLUMNAS_SISTEMA_FIN)
    return {nombre: i for i, nombre in enumerate(nombres)}


def _escribir_leyenda(ws, f):
    for i, clave in enumerate(COLORES):
        fila, col = i % 2, (i // 2) * 2
        ws.write(fila, col, COLORES[clave][2], f[f"leyenda_{clave}"])
        ws.write(fila, col + 1, COLORES[clave][3])


def escribir_hoja_programa(wb, f, programa, configuracion):
    """Hoja de un programa: leyenda, encabezados por color, validaciones y fórmula de la columna Programa."""
    ws = wb.add_worksheet(programa)
    filas = configuracion.filas
    ultima = PRIMERA_FILA_DATOS + filas - 1
    idx = indices(configuracion)
    _escribir_leyenda(ws, f)
    ws.write(0, 5, f"Programa: {programa}", f["titulo"])
    ws.write(1, 5, "Si no puede editar una celda, no corresponde a su rol o la fila es de otra persona.", f["nota"])
    ws.set_row(FILA_ENCABEZADO, 45)

    # Columna Programa (sistema): fórmula que se llena sola al registrar un curso.
    ws.write(FILA_ENCABEZADO, 0, COLUMNA_PROGRAMA, f["enc_admin"])
    ws.set_column(0, 0, 14)
    letras_clave = [xl_col_to_name(idx[c.nombre]) for c in configuracion.columnas if c.identifica]
    for r in range(PRIMERA_FILA_DATOS, ultima + 1):
        cond = "&".join(f"{l}{r + 1}" for l in letras_clave)
        ws.write_formula(r, 0, f'=IF({cond}="","","{programa}")', f["sistema"], "")

    for col in configuracion.columnas:
        c = idx[col.nombre]
        color = col.color()
        ws.write(FILA_ENCABEZADO, c, col.nombre, f[f"enc_{color}"])
        ws.set_column(c, c, col.ancho)
        formato = f[f"fecha_{color}"] if col.tipo == "fecha" else f[f"dato_{color}"]
        for r in range(PRIMERA_FILA_DATOS, ultima + 1):
            ws.write_blank(r, c, None, formato)
        _validacion(ws, col, c, ultima)

    for nombre in COLUMNAS_SISTEMA_FIN:
        c = idx[nombre]
        ws.write(FILA_ENCABEZADO, c, nombre, f["enc_admin"])
        ws.set_column(c, c, 30)
        formato = f["sistema_fecha"] if nombre.lower().startswith("fecha") else f["sistema"]
        for r in range(PRIMERA_FILA_DATOS, ultima + 1):
            ws.write_blank(r, c, None, formato)

    ws.autofilter(FILA_ENCABEZADO, 0, ultima, len(idx) - 1)
    ws.freeze_panes(PRIMERA_FILA_DATOS, idx["Profesor"] + 1 if "Profesor" in idx else 2)
    return ws


def _validacion(ws, col, c, ultima):
    opciones = opciones_de(col)
    if opciones:
        ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {"validate": "list", "source": opciones})
    elif col.tipo == "fecha":
        ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {
            "validate": "date", "criteria": ">=", "value": "2020-01-01",
            "input_title": "Fecha", "input_message": "Formato dd/mm/aaaa",
            "error_message": "Ingrese una fecha válida (dd/mm/aaaa).",
        })
    elif col.tipo == "numero":
        ws.data_validation(PRIMERA_FILA_DATOS, c, ultima, c, {
            "validate": "integer", "criteria": ">=", "value": 0,
            "error_message": "Ingrese un número entero mayor o igual a 0.",
        })


def escribir_encabezados_consolidado(wb, f, configuracion):
    ws = wb.add_worksheet("Consolidado")
    ws.write(0, 0, "Consolidado de todos los programas (solo lectura)", f["titulo"])
    ws.set_row(FILA_ENCABEZADO, 45)
    ws.write(FILA_ENCABEZADO, 0, COLUMNA_PROGRAMA, f["enc_admin"])
    ws.set_column(0, 0, 14)
    idx = indices(configuracion)
    for col in configuracion.columnas:
        c = idx[col.nombre]
        ws.write(FILA_ENCABEZADO, c, col.nombre, f[f"enc_{col.color()}"])
        ws.set_column(c, c, col.ancho, f["fecha"] if col.tipo == "fecha" else None)
    for nombre in COLUMNAS_SISTEMA_FIN:
        c = idx[nombre]
        ws.write(FILA_ENCABEZADO, c, nombre, f["enc_admin"])
        ws.set_column(c, c, 30, f["fecha_hora"] if nombre.lower().startswith("fecha") else None)
    ws.freeze_panes(PRIMERA_FILA_DATOS, idx["Profesor"] + 1 if "Profesor" in idx else 2)
    return ws


def escribir_hoja_permisos(wb, f, configuracion):
    """Hoja «Permisos»: quién edita cada columna. El Apps Script la lee para aplicar las protecciones."""
    ws = wb.add_worksheet("Permisos")
    encabezados = ["Columna", "Profesores", "Gestores", "CASA", "Identifica el curso", "Tipo"]
    for c, texto in enumerate(encabezados):
        ws.write(0, c, texto, f["enc_tabla"])
    for r, col in enumerate(configuracion.columnas, start=1):
        valores = [col.nombre] + ["SI" if v else "NO" for v in (col.profesores, col.gestores, col.casa, col.identifica)]
        valores.append(TIPOS[col.tipo] + (f": {', '.join(opciones_de(col))}" if col.tipo == "lista" else ""))
        for c, valor in enumerate(valores):
            ws.write(r, c, valor, f["celda_tabla"])
    ultima = len(configuracion.columnas)
    ws.data_validation(1, 1, ultima, 4, {"validate": "list", "source": ["SI", "NO"]})
    ws.set_column(0, 0, 48)
    ws.set_column(1, 4, 13)
    ws.set_column(5, 5, 34)
    ws.set_column(7, 7, 80)
    ws.freeze_panes(1, 1)
    ayuda = [
        "Quién puede editar cada columna en todas las hojas de programa. ADMIN siempre puede.",
        "Profesores y Gestores se refieren a los del programa de cada hoja (hoja «Usuarios»).",
        "«Identifica el curso»: la fila cuenta como registrada cuando alguna de estas columnas tiene datos.",
        "Para cambiar permisos con el libro en uso: edite SI/NO aquí (solo ADMIN) y luego",
        "menú CASA → Aplicar permisos ahora. No cambie los nombres de las columnas.",
        "Para agregar o quitar columnas, genere un libro nuevo desde la app.",
    ]
    ws.write(0, 7, "Cómo usar esta hoja", f["enc_tabla"])
    for r, texto in enumerate(ayuda, start=1):
        ws.write(r, 7, texto, f["texto"])
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
            shutil.move(str(origen), str(destino))  # otro disco o unidad
        except OSError as error:
            raise ErrorGeneracion(f"No se pudo guardar {destino.name} en {destino.parent}. "
                                  f"¿Está abierto en otro programa o la carpeta es de solo lectura? ({error})") from error
