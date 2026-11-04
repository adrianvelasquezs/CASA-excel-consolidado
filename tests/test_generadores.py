import shutil

import openpyxl
import pytest

from cronograma_neon import cli
from cronograma_neon.comun.plantillas import (
    AZUL_RAW,
    COLUMNAS_ADICIONALES,
    ErrorGeneracion,
    leer_plantillas,
    unificar_columnas,
)
from cronograma_neon.google_sheets import generador as google

NUEVAS = [n for n, _, _ in COLUMNAS_ADICIONALES]


def test_lee_plantillas_desde_carpeta_o_archivos(plantillas):
    desde_carpeta, _ = leer_plantillas([plantillas])
    assert len(desde_carpeta) == 12
    dos = sorted(plantillas.glob("*.xlsx"))[:2]
    desde_archivos, _ = leer_plantillas(dos)
    assert list(desde_archivos) == [p.name.split(" ")[0] for p in dos]


def test_rechaza_programa_repetido(plantillas, tmp_path):
    original = next(plantillas.glob("AFIN*.xlsx"))
    copia = tmp_path / "AFIN copia.xlsx"
    shutil.copy(original, copia)
    with pytest.raises(ErrorGeneracion, match="dos plantillas"):
        leer_plantillas([original, copia])


def test_rechaza_ruta_inexistente(tmp_path):
    with pytest.raises(ErrorGeneracion, match="No existe"):
        leer_plantillas([tmp_path / "no-existe"])


def test_columnas_nuevas_al_final_y_verdes(plantillas):
    columnas, avisos = unificar_columnas(leer_plantillas([plantillas])[0])
    assert [n for n, _ in columnas[-3:]] == NUEVAS
    assert all(not es_azul for _, es_azul in columnas[-3:])  # verdes: solo CASA
    assert len(columnas) == 35
    assert not any(n in aviso for aviso in avisos for n in NUEVAS)  # no se reportan como faltantes


def test_columna_nueva_que_ya_trae_una_plantilla_no_se_duplica(plantillas, tmp_path):
    original = next(plantillas.glob("AFIN*.xlsx"))
    libro = openpyxl.load_workbook(original)
    hoja = libro.worksheets[0]
    hoja.cell(3, hoja.max_column + 1, NUEVAS[0])
    copia = tmp_path / "AFIN con reporte.xlsx"
    libro.save(copia)
    columnas, _ = unificar_columnas(leer_plantillas([copia])[0])
    assert [n for n, _ in columnas].count(NUEVAS[0]) == 1


def test_plantilla_google(plantillas, tmp_path):
    destino = tmp_path / "salida" / "plantilla.xlsx"  # la carpeta de destino se crea si no existe
    google.generar_plantilla([plantillas], destino, 20, {"DOMINIO": "ejemplo.edu.co"})
    libro = openpyxl.load_workbook(destino)
    assert libro.sheetnames[0] == "Instrucciones" and libro.sheetnames[-2:] == ["Consolidado", "Usuarios"]
    assert "@ejemplo.edu.co" in libro["Usuarios"]["F2"].value

    hoja = libro["AFIN"]
    encabezados = [c.value for c in hoja[3]]
    assert encabezados[-5:] == NUEVAS + list(google.COLUMNAS_SISTEMA)
    for celda in hoja[3][-5:-2]:  # las nuevas tienen encabezado verde, no azul
        assert celda.fill.fgColor.rgb != AZUL_RAW
    col_reporte = encabezados.index("¿Se entregó reporte?") + 1
    listas = [dv for dv in hoja.data_validations.dataValidation if dv.type == "list"]
    reporte = next(dv for dv in listas if openpyxl.utils.get_column_letter(col_reporte) + "4" in dv.sqref)
    assert reporte.formula1 == '"SI,NO"'
    assert [c.value for c in libro["Consolidado"][3]] == encabezados


def test_config_apps_script():
    por_defecto = google.config_apps_script_por_defecto()
    assert google.codigo_apps_script() == google.ARCHIVO_CODIGO.read_text(encoding="utf-8")
    codigo = google.codigo_apps_script({"DOMINIO": "@Otro.edu.co", "HORA_RESTAURACION": 5, "ALERTAS_POR_CORREO": False})
    assert "DOMINIO: 'otro.edu.co'," in codigo and "HORA_RESTAURACION: 5," in codigo and "ALERTAS_POR_CORREO: false," in codigo
    assert google.config_apps_script_por_defecto() == por_defecto  # el archivo incluido no cambia
    with pytest.raises(ErrorGeneracion, match="Dominio"):
        google.codigo_apps_script({"DOMINIO": "no valido"})
    with pytest.raises(ErrorGeneracion, match="hora"):
        google.codigo_apps_script({"HORA_RESTAURACION": 24})


def test_cli_google(plantillas, tmp_path):
    assert cli.main(["google", "--plantillas", str(plantillas), "--salida", str(tmp_path / "gs.xlsx"),
                     "--apps-script", str(tmp_path / "as"), "--dominio", "ejemplo.edu.co", "--filas", "20"]) == 0
    assert (tmp_path / "gs.xlsx").exists()
    assert "DOMINIO: 'ejemplo.edu.co'," in (tmp_path / "as" / "Codigo.gs").read_text(encoding="utf-8")
    assert (tmp_path / "as" / "appsscript.json").exists()


def test_cli_ya_no_tiene_onedrive(capsys):
    with pytest.raises(SystemExit):
        cli.main(["onedrive", "--help"])
    assert "invalid choice" in capsys.readouterr().err


def test_cli_error_legible(tmp_path, capsys):
    assert cli.main(["google", "--plantillas", str(tmp_path / "nada"), "--salida", str(tmp_path / "x.xlsx")]) == 1
    assert "Error: No existe" in capsys.readouterr().err


def test_verificar(tmp_path):
    informe = tmp_path / "informe.txt"
    assert cli.main(["verificar", "--informe", str(informe)]) == 0
    assert "OK:" in informe.read_text(encoding="utf-8")
