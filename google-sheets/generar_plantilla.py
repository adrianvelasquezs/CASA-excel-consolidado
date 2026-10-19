"""
Genera la plantilla para Google Sheets del Cronograma 2027 Bloque Neón a partir
de raw/ (que nunca se modifica).

Entradas:
  raw/*.xlsx   plantillas por programa

Resultado:
  google-sheets/salida/Cronograma 2027 Bloque Neón - Google Sheets.xlsx

El archivo NO lleva contraseñas ni protecciones de Excel: se importa a Google
Sheets y los permisos los aplica el Apps Script de google-sheets/apps-script/
según los correos registrados en la hoja «Usuarios».

Uso (desde la raíz del proyecto):
  uv run google-sheets/generar_plantilla.py [--filas 150]
"""

import argparse
import sys
from pathlib import Path

import xlsxwriter

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI.parent / "comun"))

from plantillas import (  # noqa: E402
    ROOT,
    crear_formatos,
    escribir_base_programa,
    escribir_encabezados_consolidado,
    escribir_lineas,
    leer_plantillas,
    unificar_columnas,
)

SALIDA_DIR = AQUI / "salida"
ARCHIVO_SALIDA = SALIDA_DIR / "Cronograma 2027 Bloque Neón - Google Sheets.xlsx"

# Deben coincidir con CONFIG en google-sheets/apps-script/Codigo.gs
COLUMNAS_SISTEMA = ("Registrado por", "Fecha de registro")
ROLES = ["ADMIN", "CASA", "GESTOR", "PROFESOR"]
FILAS_USUARIOS = 500


def escribir_usuarios(wb, f, programas):
    ws = wb.add_worksheet("Usuarios")
    encabezados = ["Correo", "Rol", "Programa(s)", "Nombre (opcional)"]
    for c, texto in enumerate(encabezados):
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
        "input_message": "Para varios programas sepárelos con coma: AFIN, MF",
        "show_error": False,  # permite escribir varios programas separados por coma
    })
    ayuda = [
        "Una fila por persona. Use el correo institucional @uniandes.edu.co con el que entra a Google.",
        "ADMIN: todo el libro (administra permisos). No necesita programa.",
        "CASA: columnas verdes de todos los programas. No necesita programa.",
        "GESTOR: columnas azules de su(s) programa(s), incluidas las filas de sus profesores.",
        "PROFESOR: columnas azules de su(s) programa(s); las filas que registre quedan solo para quien las registró.",
        "Después de modificar esta hoja: menú CASA → Aplicar permisos ahora.",
    ]
    ws.write(0, 5, "Cómo llenar esta hoja", f["enc_tabla"])
    for r, texto in enumerate(ayuda, start=1):
        ws.write(r, 5, texto, f["texto"])
    return ws


def escribir_instrucciones(wb, f, programas):
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
        ("texto", "1. Entre a Google con su cuenta institucional @uniandes.edu.co (no con una cuenta personal de Gmail)."),
        ("texto", "2. Abra la hoja de su programa y registre cada curso en una fila vacía."),
        ("texto", "3. Si Google muestra «No tienes permiso para editar», esa celda no le corresponde a su rol "
                  "o la fila es de otro profesor."),
        ("texto", "4. Para liberar una fila que registró por error, borre todo su contenido azul: "
                  "la fila vuelve a quedar disponible."),
        ("texto", "5. Las fechas van en formato dd/mm/aaaa; las columnas SI/NO tienen lista desplegable."),
        ("h2", "Programas incluidos"),
        ("texto", ", ".join(programas)),
    ])
    return ws


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--filas", type=int, default=150, help="filas de registro por programa (default 150)")
    args = parser.parse_args()

    plantillas, anchos = leer_plantillas()
    programas = list(plantillas)
    columnas = unificar_columnas(plantillas)

    SALIDA_DIR.mkdir(exist_ok=True)
    wb = xlsxwriter.Workbook(str(ARCHIVO_SALIDA))
    f = crear_formatos(wb)
    escribir_instrucciones(wb, f, programas)
    for p in programas:
        escribir_base_programa(
            wb, f, p, columnas, anchos, args.filas,
            "Gris = lo llena el sistema. Si no puede editar una celda, no corresponde a su rol.",
            COLUMNAS_SISTEMA,
        )
    consolidado = escribir_encabezados_consolidado(wb, f, columnas, anchos, COLUMNAS_SISTEMA)
    consolidado.write(1, 0, "La fórmula de esta hoja la escribe el script (menú CASA → Configurar libro).", f["nota"])
    escribir_usuarios(wb, f, programas)
    wb.close()

    print(f"Generado: {ARCHIVO_SALIDA.relative_to(ROOT)}")
    print(f"{len(programas)} programas, {len(columnas)} columnas + {len(COLUMNAS_SISTEMA)} de sistema, "
          f"{args.filas} filas por programa")


if __name__ == "__main__":
    main()
