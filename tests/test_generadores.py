import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from cronograma_neon import cli
from cronograma_neon.comun.configuracion import Columna, Configuracion, ErrorGeneracion, predeterminada
from cronograma_neon.comun.estructura import COLORES
from cronograma_neon.google_sheets import generador as google

NUEVAS = ["¿Se entregó reporte?", "Razones de porqué están pendientes reportes", "Razones de porqué no se realizó medición"]


def color(celda):
    return celda.fill.fgColor.rgb[-6:].upper() if celda.fill.fill_type else None


def test_predeterminada_conserva_el_cronograma_original():
    config = predeterminada()
    assert len(config.programas) == 12 and config.programas[0] == "AFIN"
    nombres = [c.nombre for c in config.columnas]
    assert len(nombres) == 34 and nombres[-3:] == NUEVAS
    assert [c.nombre for c in config.columnas if c.identifica] == ["Código de curso", "CRN", "Curso"]
    assert all(c.casa and not c.es_del_programa for c in config.columnas[-3:])  # columnas de reporte: solo CASA
    curso = next(c for c in config.columnas if c.nombre == "Curso")
    assert curso.profesores and curso.gestores and not curso.casa


@pytest.mark.parametrize("cambio, mensaje", [
    (lambda c: c.programas.append("AAA"), "repetido"),
    (lambda c: c.programas.append("Consolidado"), "reservado"),
    (lambda c: c.programas.append("a/b"), "nombre de hoja"),
    (lambda c: c.columnas.append(Columna("curso")), "repetida"),
    (lambda c: c.columnas.append(Columna("Registrado por")), "sistema"),
    (lambda c: c.columnas.append(Columna("Lista", tipo="lista", opciones=["Una"])), "dos opciones"),
    (lambda c: c.columnas.append(Columna("Lista", tipo="lista", opciones=["a,b", "c"])), "comas"),
    (lambda c: c.columnas.append(Columna("Ancha", ancho=500)), "ancho"),
    (lambda c: [setattr(col, "identifica", False) for col in c.columnas], "Identifica el curso"),
    (lambda c: c.columnas.clear(), "al menos una columna"),
    (lambda c: setattr(c, "filas", 5), "filas"),
])
def test_validaciones(pequena, cambio, mensaje):
    cambio(pequena)
    with pytest.raises(ErrorGeneracion, match=mensaje):
        pequena.validar()


def test_json_ida_y_vuelta(pequena, tmp_path):
    ruta = tmp_path / "config.json"
    pequena.guardar_json(ruta)
    assert Configuracion.leer_json(ruta) == pequena
    ruta.write_text("{no es json", encoding="utf-8")
    with pytest.raises(ErrorGeneracion, match="JSON"):
        Configuracion.leer_json(ruta)


def test_genera_desde_cero_con_columnas_y_permisos(pequena, tmp_path):
    destino = tmp_path / "salida" / "plantilla.xlsx"  # la carpeta de destino se crea si no existe
    google.generar_plantilla(pequena, destino, {"DOMINIO": "ejemplo.edu.co"})
    libro = openpyxl.load_workbook(destino)
    assert libro.sheetnames == ["Instrucciones", "AAA", "BBB", "Consolidado", "Permisos", "Usuarios"]

    hoja = libro["AAA"]
    encabezados = [c.value for c in hoja[3]]
    assert encabezados == ["Programa", "Curso", "Inicio", "Nivel", "Estudiantes", "¿Revisado?", "Nota interna",
                           "Registrado por", "Fecha de registro"]
    colores = [color(c) for c in hoja[3]]
    assert colores[1] == COLORES["programa"][0][1:]    # Curso: programa (azul)
    assert colores[3] == COLORES["programa"][0][1:]    # Nivel: solo gestores, también del programa
    assert colores[4] == COLORES["mixto"][0][1:]       # Estudiantes: programa y CASA
    assert colores[5] == COLORES["casa"][0][1:]        # ¿Revisado?: CASA
    assert colores[6] == COLORES["admin"][0][1:]       # Nota interna: solo ADMIN
    assert hoja["A4"].value == '=IF(B4="","","AAA")'    # solo «Curso» identifica el curso
    assert hoja.column_dimensions["B"].width == pytest.approx(30, abs=1)

    listas = {str(dv.sqref).split(":")[0]: dv.formula1 for dv in hoja.data_validations.dataValidation if dv.type == "list"}
    assert listas["D4"] == '"Bajo,Medio,Alto"' and listas["F4"] == '"SI,NO"'
    tipos = {str(dv.sqref).split(":")[0]: dv.type for dv in hoja.data_validations.dataValidation}
    assert tipos["C4"] == "date" and tipos["E4"] == "whole"

    permisos = [list(f[:5]) for f in libro["Permisos"].iter_rows(min_row=1, max_row=7, values_only=True)]
    assert permisos[0] == ["Columna", "Profesores", "Gestores", "CASA", "Identifica el curso"]
    assert permisos[3] == ["Nivel", "NO", "SI", "NO", "NO"]
    assert permisos[6] == ["Nota interna", "NO", "NO", "NO", "NO"]
    assert [c.value for c in libro["Consolidado"][3]] == encabezados
    assert "@ejemplo.edu.co" in libro["Usuarios"]["F2"].value


def test_predeterminada_genera_las_columnas_nuevas_en_verde(tmp_path):
    destino = tmp_path / "p.xlsx"
    google.generar_plantilla(predeterminada(), destino)
    hoja = openpyxl.load_workbook(destino)["MF"]
    encabezados = [c.value for c in hoja[3]]
    assert encabezados[-5:] == NUEVAS + ["Registrado por", "Fecha de registro"]
    assert all(color(c) == COLORES["casa"][0][1:] for c in hoja[3][-5:-2])
    reporte = get_column_letter(encabezados.index("¿Se entregó reporte?") + 1)
    assert any(str(dv.sqref).startswith(reporte + "4") and dv.formula1 == '"SI,NO"'
               for dv in hoja.data_validations.dataValidation)


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


def test_cli_con_configuracion_propia(pequena, tmp_path):
    pequena.guardar_json(tmp_path / "c.json")
    assert cli.main(["google", "--config", str(tmp_path / "c.json"), "--salida", str(tmp_path / "gs.xlsx"),
                     "--filas", "15", "--apps-script", str(tmp_path / "as"), "--dominio", "ejemplo.edu.co"]) == 0
    libro = openpyxl.load_workbook(tmp_path / "gs.xlsx")
    assert libro.sheetnames[1:3] == ["AAA", "BBB"]
    assert "DOMINIO: 'ejemplo.edu.co'," in (tmp_path / "as" / "Codigo.gs").read_text(encoding="utf-8")


def test_cli_exportar_y_usar_predeterminada(tmp_path):
    assert cli.main(["exportar-configuracion", "--salida", str(tmp_path / "pred.json")]) == 0
    assert Configuracion.leer_json(tmp_path / "pred.json") == predeterminada()
    assert cli.main(["google", "--salida", str(tmp_path / "gs.xlsx")]) == 0


def test_cli_errores_legibles(tmp_path, capsys):
    (tmp_path / "malo.json").write_text('{"programas": [], "columnas": []}', encoding="utf-8")
    assert cli.main(["google", "--config", str(tmp_path / "malo.json"), "--salida", str(tmp_path / "x.xlsx")]) == 1
    assert "Error: Agregue al menos un programa" in capsys.readouterr().err
    assert not (tmp_path / "x.xlsx").exists()
    with pytest.raises(SystemExit):
        cli.main(["onedrive", "--help"])


def test_verificar(tmp_path):
    informe = tmp_path / "informe.txt"
    assert cli.main(["verificar", "--informe", str(informe)]) == 0
    assert "OK:" in informe.read_text(encoding="utf-8")
