"""
Línea de comandos. Todas las rutas se indican explícitamente.

  cronograma-neon google --plantillas CARPETA_O_ARCHIVOS... --salida PLANTILLA.xlsx \
                         [--apps-script CARPETA] [--dominio uniandes.edu.co] [--filas 150]
  cronograma-neon verificar [--informe INFORME.txt]

Sin argumentos se abre la interfaz gráfica.
"""

import argparse
import sys
import tempfile
import traceback
from pathlib import Path

from . import __version__
from .comun.plantillas import COLUMNAS_ADICIONALES, ErrorGeneracion, carpeta_plantillas_incluidas
from .google_sheets import generador as google


def _google(args):
    config = {}
    if args.dominio:
        config["DOMINIO"] = args.dominio
    if args.hora is not None:
        config["HORA_RESTAURACION"] = args.hora
    if args.sin_alertas:
        config["ALERTAS_POR_CORREO"] = False
    if args.gestores_sin_filas_de_profesores:
        config["GESTOR_EDITA_FILAS_DE_PROFESORES"] = False
    mensajes = google.generar_plantilla(args.plantillas, args.salida, args.filas, config)
    if args.apps_script:
        carpeta = Path(args.apps_script)
        carpeta.mkdir(parents=True, exist_ok=True)
        google.guardar_texto(carpeta / "Codigo.gs", google.codigo_apps_script(config))
        google.guardar_texto(carpeta / "appsscript.json", google.manifiesto_apps_script())
        mensajes.append(f"Apps Script guardado en: {carpeta}")
    return mensajes


def verificar():
    """Prueba de humo del ejecutable: genera la plantilla con las plantillas incluidas en una carpeta temporal."""
    import openpyxl

    mensajes = [f"Cronograma Neón {__version__} — verificación"]
    incluidas = carpeta_plantillas_incluidas()
    if not incluidas:
        raise ErrorGeneracion("La app no trae plantillas incluidas.")
    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / "google.xlsx"
        mensajes += google.generar_plantilla([incluidas], destino, 20)
        libro = openpyxl.load_workbook(destino)
        if libro.sheetnames[0] != "Instrucciones" or libro.sheetnames[-2:] != ["Consolidado", "Usuarios"]:
            raise ErrorGeneracion(f"Hojas inesperadas en la plantilla: {libro.sheetnames}")
        encabezados = [c.value for c in libro[libro.sheetnames[1]][3]]
        faltan = [n for n, _, _ in COLUMNAS_ADICIONALES if n not in encabezados]
        if faltan:
            raise ErrorGeneracion(f"Faltan columnas en la plantilla: {faltan}")
        codigo = google.codigo_apps_script({"DOMINIO": "ejemplo.edu.co"})
        if "DOMINIO: 'ejemplo.edu.co'," not in codigo:
            raise ErrorGeneracion("No se aplicó la configuración del Apps Script.")
        mensajes.append(f"OK: {len(libro.sheetnames) - 3} programas, {len(encabezados)} columnas, "
                        f"Apps Script de {len(codigo)} caracteres.")
    return mensajes


def construir_parser():
    parser = argparse.ArgumentParser(prog="cronograma-neon", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="comando", required=True)

    gs = sub.add_parser("google", help="Plantilla para Google Sheets (+ Apps Script)")
    gs.add_argument("--plantillas", nargs="+", required=True, help="carpetas y/o archivos .xlsx de las plantillas")
    gs.add_argument("--salida", required=True, help="dónde guardar la plantilla generada")
    gs.add_argument("--apps-script", help="carpeta donde guardar Codigo.gs y appsscript.json (opcional)")
    gs.add_argument("--dominio", help="dominio institucional (por defecto el de Codigo.gs)")
    gs.add_argument("--hora", type=int, help="hora diaria de restauración de permisos (0-23)")
    gs.add_argument("--sin-alertas", action="store_true", help="no enviar alertas por correo a los ADMIN")
    gs.add_argument("--gestores-sin-filas-de-profesores", action="store_true",
                    help="los gestores no pueden editar las filas registradas por profesores")
    gs.add_argument("--filas", type=int, default=150, help="filas de registro por programa (por defecto 150)")

    ver = sub.add_parser("verificar", help="prueba rápida de la instalación (la usa el CI)")
    ver.add_argument("--informe", help="archivo donde escribir el resultado")
    return parser


def main(argv=None):
    for flujo in (sys.stdout, sys.stderr):  # consolas de Windows sin UTF-8: no fallar por un acento
        if hasattr(flujo, "reconfigure"):
            flujo.reconfigure(errors="replace")
    args = construir_parser().parse_args(argv)
    informe = getattr(args, "informe", None)
    try:
        mensajes = _google(args) if args.comando == "google" else verificar()
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
