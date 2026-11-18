"""Corre la simulación de Google Sheets (Node) contra la plantilla generada y el Codigo.gs incluido."""

import json
import os
import shutil
import subprocess
from pathlib import Path

import openpyxl
import pytest

from cronograma_neon.comun.configuracion import predeterminada
from cronograma_neon.google_sheets import generador as google

SIMULACION = Path(__file__).resolve().parent / "apps_script" / "simulacion.js"


@pytest.mark.skipif(shutil.which("node") is None, reason="requiere Node.js")
@pytest.mark.parametrize("sin_permisos", [False, True], ids=["con_hoja_permisos", "sin_hoja_permisos"])
def test_simulacion_apps_script(tmp_path, sin_permisos):
    destino = tmp_path / "plantilla.xlsx"
    google.generar_plantilla(predeterminada(), destino)
    estructura = {}
    for ws in openpyxl.load_workbook(destino).worksheets:
        estructura[ws.title] = {
            "header": [(c.value, c.fill.fgColor.rgb if c.fill.fill_type else None) for c in ws[3]],
            "max_row": ws.max_row,
            "usuarios_hdr": [c.value for c in ws[1]][:4],
        }
        if ws.title == "Permisos" and sin_permisos:  # libro de una versión anterior, sin la hoja
            del estructura[ws.title]
        elif ws.title == "Permisos":
            estructura[ws.title]["filas"] = [["" if v is None else v for v in fila[:5]] for fila in ws.iter_rows(values_only=True)]
    archivo = tmp_path / "estructura.json"
    archivo.write_text(json.dumps(estructura, ensure_ascii=False), encoding="utf-8")

    entorno = {**os.environ, "SIN_PERMISOS": "1"} if sin_permisos else None
    resultado = subprocess.run(["node", str(SIMULACION), str(google.ARCHIVO_CODIGO), str(archivo)],
                               capture_output=True, text=True, encoding="utf-8", env=entorno)
    assert resultado.returncode == 0, resultado.stdout + resultado.stderr
    assert "Todas las pruebas pasaron" in resultado.stdout
