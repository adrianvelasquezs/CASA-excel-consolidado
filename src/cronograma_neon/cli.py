"""
Línea de comandos. Todas las rutas se indican explícitamente.

  cronograma-neon google --salida PLANTILLA.xlsx [--config CONFIGURACION.json] [--filas 150] \
                         [--apps-script CARPETA] [--dominio uniandes.edu.co]
  cronograma-neon exportar-configuracion --salida CONFIGURACION.json
  cronograma-neon verificar [--informe INFORME.txt]

Sin --config se usa la configuración predeterminada (programas, columnas y permisos).
Sin argumentos se abre la interfaz gráfica.
"""

import argparse
import sys
import tempfile
import traceback
from pathlib import Path

from . import __version__
from .comun.configuracion import Configuracion, ErrorGeneracion, predeterminada
from .google_sheets import generador as google


def _google(args):
    configuracion = Configuracion.leer_json(args.config) if args.config else predeterminada()
    if args.filas is not None:
        configuracion.filas = args.filas
    config = {}
    if args.dominio:
        config["DOMINIO"] = args.dominio
    if args.hora is not None:
        config["HORA_RESTAURACION"] = args.hora
    if args.sin_alertas:
        config["ALERTAS_POR_CORREO"] = False
    if args.gestores_sin_filas_de_profesores:
        config["GESTOR_EDITA_FILAS_DE_PROFESORES"] = False
    mensajes = google.generar_plantilla(configuracion, args.salida, config)
    if args.apps_script:
        carpeta = Path(args.apps_script)
        carpeta.mkdir(parents=True, exist_ok=True)
        google.guardar_texto(carpeta / "Codigo.gs", google.codigo_apps_script(config))
        google.guardar_texto(carpeta / "appsscript.json", google.manifiesto_apps_script())
        mensajes.append(f"Apps Script guardado en: {carpeta}")
    return mensajes


def _exportar(args):
    predeterminada().guardar_json(args.salida)
    return [f"Configuración predeterminada guardada en: {args.salida}"]


def verificar():
    """Prueba de humo del ejecutable: genera la plantilla con la configuración predeterminada en una carpeta temporal."""
    import openpyxl

    mensajes = [f"Cronograma Neón {__version__} — verificación"]
    configuracion = predeterminada()
    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / "google.xlsx"
        mensajes += google.generar_plantilla(configuracion, destino)
        libro = openpyxl.load_workbook(destino)
        esperadas = ["Instrucciones", *configuracion.programas, "Consolidado", "Permisos", "Usuarios"]
        if libro.sheetnames != esperadas:
            raise ErrorGeneracion(f"Hojas inesperadas en la plantilla: {libro.sheetnames}")
        encabezados = [c.value for c in libro[configuracion.programas[0]][3]]
        faltan = [c.nombre for c in configuracion.columnas if c.nombre not in encabezados]
        if faltan:
            raise ErrorGeneracion(f"Faltan columnas en la plantilla: {faltan}")
        if libro["Permisos"].max_row != len(configuracion.columnas) + 1:
            raise ErrorGeneracion("La hoja Permisos no tiene una fila por columna.")
        codigo = google.codigo_apps_script({"DOMINIO": "ejemplo.edu.co"})
        if "DOMINIO: 'ejemplo.edu.co'," not in codigo:
            raise ErrorGeneracion("No se aplicó la configuración del Apps Script.")
        mensajes.append(f"OK: {len(configuracion.programas)} programas, {len(encabezados)} columnas, "
                        f"Apps Script de {len(codigo)} caracteres.")
    return mensajes


def construir_parser():
    parser = argparse.ArgumentParser(prog="cronograma-neon", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="comando", required=True)

    gs = sub.add_parser("google", help="Plantilla para Google Sheets (+ Apps Script)")
    gs.add_argument("--salida", required=True, help="dónde guardar la plantilla generada")
    gs.add_argument("--config", help="configuración JSON (programas, columnas y permisos); por defecto la predeterminada")
    gs.add_argument("--filas", type=int, help="filas de registro por programa (por defecto, las de la configuración)")
    gs.add_argument("--apps-script", help="carpeta donde guardar Codigo.gs y appsscript.json (opcional)")
    gs.add_argument("--dominio", help="dominio institucional (por defecto el de Codigo.gs)")
    gs.add_argument("--hora", type=int, help="hora diaria de restauración de permisos (0-23)")
    gs.add_argument("--sin-alertas", action="store_true", help="no enviar alertas por correo a los ADMIN")
    gs.add_argument("--gestores-sin-filas-de-profesores", action="store_true",
                    help="los gestores no pueden editar las filas registradas por profesores")

    ex = sub.add_parser("exportar-configuracion", help="guarda la configuración predeterminada como JSON para editarla")
    ex.add_argument("--salida", required=True, help="dónde guardar el JSON")

    ver = sub.add_parser("verificar", help="prueba rápida de la instalación (la usa el CI)")
    ver.add_argument("--informe", help="archivo donde escribir el resultado")
    return parser


def main(argv=None):
    for flujo in (sys.stdout, sys.stderr):  # consolas de Windows sin UTF-8: no fallar por un acento
        if hasattr(flujo, "reconfigure"):
            flujo.reconfigure(errors="replace")
    args = construir_parser().parse_args(argv)
    informe = getattr(args, "informe", None)
    acciones = {"google": _google, "exportar-configuracion": _exportar, "verificar": lambda _a: verificar()}
    try:
        mensajes = acciones[args.comando](args)
        codigo = 0
    except ErrorGeneracion as error:
        mensajes, codigo = [f"Error: {error}"], 1
    except Exception:  # noqa: BLE001 - en la verificación se quiere el detalle completo
        mensajes, codigo = ["Error inesperado:", traceback.format_exc()], 2
    texto = "\n".join(mensajes)
    print(texto, file=sys.stdout if codigo == 0 else sys.stderr)
    if informe:
        Path(informe).write_text(texto + "\n", encoding="utf-8")
    return codigo
