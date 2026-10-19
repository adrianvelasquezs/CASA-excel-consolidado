"""
Versión Google Sheets: plantilla para importar a Google Sheets + Apps Script de permisos.

La plantilla NO lleva contraseñas ni protecciones de Excel: los permisos los
aplica el Apps Script (apps_script/Codigo.gs) según los correos de la hoja
«Usuarios». Las opciones del script (dominio, alertas, etc.) se aplican al
exportar el código; el archivo del paquete no se modifica.
"""

import re
import tempfile
from pathlib import Path

import xlsxwriter

from ..comun.plantillas import (
    ErrorGeneracion,
    crear_formatos,
    escribir_base_programa,
    escribir_encabezados_consolidado,
    escribir_lineas,
    leer_plantillas,
    reemplazar_archivo,
    unificar_columnas,
    validar_filas,
)

NOMBRE_SUGERIDO = "Cronograma 2027 Bloque Neón - Google Sheets.xlsx"
CARPETA_APPS_SCRIPT = Path(__file__).resolve().parent / "apps_script"
ARCHIVO_CODIGO = CARPETA_APPS_SCRIPT / "Codigo.gs"
ARCHIVO_MANIFIESTO = CARPETA_APPS_SCRIPT / "appsscript.json"

# Deben coincidir con CONFIG en apps_script/Codigo.gs
COLUMNAS_SISTEMA = ("Registrado por", "Fecha de registro")
ROLES = ["ADMIN", "CASA", "GESTOR", "PROFESOR"]
FILAS_USUARIOS = 500

# Opciones de CONFIG en Codigo.gs que se pueden cambiar desde la interfaz o la línea de comandos.
CAMPOS_APPS_SCRIPT = {
    "DOMINIO": "texto",
    "GESTOR_EDITA_FILAS_DE_PROFESORES": "bool",
    "ALERTAS_POR_CORREO": "bool",
    "HORA_RESTAURACION": "hora",
}

# ---------------------------------------------------------------------------
# Apps Script
# ---------------------------------------------------------------------------


def config_apps_script_por_defecto():
    """Valores de CONFIG tal como están en el Codigo.gs incluido."""
    texto = ARCHIVO_CODIGO.read_text(encoding="utf-8")
    valores = {}
    for clave, tipo in CAMPOS_APPS_SCRIPT.items():
        m = re.search(rf"^\s*{clave}:\s*(.+?),\s*$", texto, re.M)
        if not m:
            raise ErrorGeneracion(f"No se encontró {clave} en Codigo.gs")
        crudo = m.group(1).strip()
        valores[clave] = crudo.strip("'\"") if tipo == "texto" else crudo == "true" if tipo == "bool" else int(crudo)
    return valores


def validar_config_apps_script(config):
    """Completa con los valores por defecto, normaliza y valida. Devuelve la configuración final."""
    final = {**config_apps_script_por_defecto(), **{k: v for k, v in (config or {}).items() if k in CAMPOS_APPS_SCRIPT}}
    dominio = str(final["DOMINIO"]).strip().lstrip("@").lower()
    if not re.fullmatch(r"[a-z0-9-]+(\.[a-z0-9-]+)+", dominio):
        raise ErrorGeneracion(f"Dominio inválido: «{final['DOMINIO']}» (ejemplo: uniandes.edu.co).")
    try:
        hora = int(final["HORA_RESTAURACION"])
    except (TypeError, ValueError):
        hora = -1
    if not 0 <= hora <= 23:
        raise ErrorGeneracion("La hora de restauración debe estar entre 0 y 23.")
    return {
        "DOMINIO": dominio,
        "HORA_RESTAURACION": hora,
        "GESTOR_EDITA_FILAS_DE_PROFESORES": bool(final["GESTOR_EDITA_FILAS_DE_PROFESORES"]),
        "ALERTAS_POR_CORREO": bool(final["ALERTAS_POR_CORREO"]),
    }


def codigo_apps_script(config=None):
    """Texto de Codigo.gs con la configuración aplicada."""
    valores = validar_config_apps_script(config)
    texto = ARCHIVO_CODIGO.read_text(encoding="utf-8")
    for clave, tipo in CAMPOS_APPS_SCRIPT.items():
        valor = valores[clave]
        if tipo == "texto":
            literal = "'" + str(valor).replace("\\", "\\\\").replace("'", "\\'") + "'"
        elif tipo == "bool":
            literal = "true" if valor else "false"
        else:
            literal = str(int(valor))
        texto, n = re.subn(rf"^(\s*{clave}:\s*).+?(,\s*)$", lambda m: m.group(1) + literal + m.group(2),
                           texto, count=1, flags=re.M)
        if n != 1:
            raise ErrorGeneracion(f"No se pudo aplicar {clave} en Codigo.gs")
    return texto


def manifiesto_apps_script():
    return ARCHIVO_MANIFIESTO.read_text(encoding="utf-8")


def guardar_texto(destino, texto):
    """Guarda Codigo.gs / appsscript.json donde elija el usuario."""
    try:
        Path(destino).write_text(texto, encoding="utf-8")
    except OSError as error:
        raise ErrorGeneracion(f"No se pudo guardar {destino}: {error}") from error


# ---------------------------------------------------------------------------
# Plantilla
# ---------------------------------------------------------------------------


def _escribir_usuarios(wb, f, programas, dominio):
    ws = wb.add_worksheet("Usuarios")
    for c, texto in enumerate(["Correo", "Rol", "Programa(s)", "Nombre (opcional)"]):
        ws.write(0, c, texto, f["enc_tabla"])
    ws.set_column(0, 0, 38)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 28)
    ws.set_column(3, 3, 32)
    ws.set_column(5, 5, 90)
    ws.freeze_panes(1, 0)
    ws.data_validation(1, 1, FILAS_USUARIOS, 1, {
        "validate": "list", "source": ROLES,
        "error_message": "Rol válido: " + ", ".join(ROLES),
    })
    ws.data_validation(1, 2, FILAS_USUARIOS, 2, {
        "validate": "list", "source": programas,
        "input_title": "Programa(s)",
        "input_message": "Para varios programas sepárelos con coma: " + ", ".join(programas[:2]),
        "show_error": False,  # permite escribir varios programas separados por coma
    })
    ayuda = [
        f"Una fila por persona. Use el correo institucional @{dominio} con el que entra a Google.",
        "ADMIN: todo el libro (administra permisos). No necesita programa.",
        "CASA: columnas verdes de todos los programas. No necesita programa.",
        "GESTOR: columnas azules de su(s) programa(s), incluidas las filas de sus profesores.",
        "PROFESOR: columnas azules de su(s) programa(s); las filas que registre quedan solo para quien las registró.",
        "Después de modificar esta hoja: menú CASA → Aplicar permisos ahora.",
    ]
    ws.write(0, 5, "Cómo llenar esta hoja", f["enc_tabla"])
    for r, texto in enumerate(ayuda, start=1):
        ws.write(r, 5, texto, f["texto"])


def _escribir_instrucciones(wb, f, programas, dominio):
    ws = wb.add_worksheet("Instrucciones")
    escribir_lineas(ws, f, [
        ("h1", "Cronograma 2027 Bloque Neón — Consolidado"),
        ("texto", "Este libro reúne en un solo archivo las plantillas de todos los programas. "
                  "Cada programa tiene su propia hoja; la hoja «Consolidado» las une automáticamente."),
        ("h2", "¿Qué puede editar cada quien?"),
        ("texto", "• Celdas AZULES: profesores y gestores del programa dueño de la hoja."),
        ("texto", "• Cuando un profesor registra un curso en una fila, esa fila queda a su nombre "
                  "(columna «Registrado por») y ningún otro profesor puede modificarla. "
                  "Los gestores del programa sí pueden."),
        ("texto", "• Celdas VERDES: solo el equipo CASA."),
        ("texto", "• Títulos, la columna «Programa», «Registrado por», «Fecha de registro», "
                  "y las hojas Consolidado, Usuarios e Instrucciones: solo administración."),
        ("h2", "¿Cómo se usa?"),
        ("texto", f"1. Entre a Google con su cuenta institucional @{dominio} (no con una cuenta personal de Gmail)."),
        ("texto", "2. Abra la hoja de su programa y registre cada curso en una fila vacía."),
        ("texto", "3. Si Google muestra «No tienes permiso para editar», esa celda no le corresponde a su rol "
                  "o la fila es de otro profesor."),
        ("texto", "4. Para liberar una fila que registró por error, borre todo su contenido azul: "
                  "la fila vuelve a quedar disponible."),
        ("texto", "5. Las fechas van en formato dd/mm/aaaa; las columnas SI/NO tienen lista desplegable."),
        ("h2", "Programas incluidos"),
        ("texto", ", ".join(programas)),
    ])


def generar_plantilla(rutas_plantillas, destino, filas=150, config_apps_script=None):
    """Genera la plantilla para Google Sheets en `destino`. Devuelve los mensajes para el usuario."""
    validar_filas(filas)
    dominio = validar_config_apps_script(config_apps_script)["DOMINIO"]
    plantillas, anchos = leer_plantillas(rutas_plantillas)
    programas = list(plantillas)
    columnas, mensajes = unificar_columnas(plantillas)

    with tempfile.TemporaryDirectory() as tmp:
        temporal = Path(tmp) / "plantilla.xlsx"
        wb = xlsxwriter.Workbook(str(temporal))
        f = crear_formatos(wb)
        _escribir_instrucciones(wb, f, programas, dominio)
        for p in programas:
            escribir_base_programa(
                wb, f, p, columnas, anchos, filas,
                "Gris = lo llena el sistema. Si no puede editar una celda, no corresponde a su rol.",
                COLUMNAS_SISTEMA,
            )
        consolidado = escribir_encabezados_consolidado(wb, f, columnas, anchos, COLUMNAS_SISTEMA)
        consolidado.write(1, 0, "La fórmula de esta hoja la escribe el script (menú CASA → Configurar libro).", f["nota"])
        _escribir_usuarios(wb, f, programas, dominio)
        wb.close()
        reemplazar_archivo(temporal, destino)

    mensajes.append(f"Plantilla guardada en: {destino}")
    mensajes.append(f"{len(programas)} programas, {len(columnas)} columnas + {len(COLUMNAS_SISTEMA)} de sistema, "
                    f"{filas} filas por programa.")
    return mensajes
