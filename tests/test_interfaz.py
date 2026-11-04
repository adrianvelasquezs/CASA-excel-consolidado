"""Maneja la interfaz por código, con los diálogos de archivos y mensajes simulados."""

import time
import tkinter as tk

import openpyxl
import pytest

from cronograma_neon.comun import ajustes
from cronograma_neon.comun.configuracion import TIPOS, Configuracion, predeterminada
from cronograma_neon.interfaz import app as interfaz


@pytest.fixture
def app(monkeypatch):
    try:
        ventana = interfaz.App()
    except tk.TclError as error:  # sin pantalla (p. ej. Linux sin X)
        pytest.skip(f"Tk no disponible: {error}")
    ventana.withdraw()
    dialogos = {"guardar": [], "abrir": [], "avisos": [], "respuestas": []}
    monkeypatch.setattr(interfaz.filedialog, "asksaveasfilename", lambda **k: dialogos["guardar"].pop(0))
    monkeypatch.setattr(interfaz.filedialog, "askopenfilename", lambda **k: dialogos["abrir"].pop(0))
    for nombre in ("showwarning", "showinfo", "showerror"):
        monkeypatch.setattr(interfaz.messagebox, nombre, lambda titulo, *_a, **_k: dialogos["avisos"].append(titulo))
    monkeypatch.setattr(interfaz.messagebox, "askyesno",
                        lambda *_a, **_k: dialogos["respuestas"].pop(0) if dialogos["respuestas"] else True)
    ventana.dialogos = dialogos
    yield ventana
    try:
        ventana.destroy()
    except tk.TclError:
        pass


def esperar(app):
    inicio = time.time()
    while app.ocupado and time.time() - inicio < 120:
        app.update()
        time.sleep(0.02)
    app.update()


def guardada():
    return Configuracion.desde_dict(ajustes.leer()["configuracion"])


def test_inicia_con_la_predeterminada(app):
    assert app.configuracion == predeterminada()
    assert len(app.columnas.tabla.get_children()) == 34
    assert "12 programas · 34 columnas" in app.estado_config.cget("text")


def test_agregar_editar_mover_y_quitar_columnas(app):
    c = app.columnas
    c.limpiar()
    c.nombre.set("Observaciones CASA"); c.casa.set(True); c.tipo.set(TIPOS["texto"]); c.ancho.set("40")
    c.agregar()
    assert app.configuracion.columnas[-1].nombre == "Observaciones CASA"
    assert guardada().columnas[-1].quien_edita() == "CASA"  # se guarda sola en las preferencias

    c.nombre.set("observaciones casa"); c.agregar()
    assert app.dialogos["avisos"][-1] == "Columna inválida"  # nombre repetido

    # Editar: «Curso» pasa a ser editable también por CASA y a lista de opciones inválida → aviso
    i = next(i for i, col in enumerate(app.configuracion.columnas) if col.nombre == "Curso")
    c.tabla.selection_set(str(i)); c.cargar_formulario()
    assert c.nombre.get() == "Curso" and c.profesores.get() and not c.casa.get()
    c.casa.set(True); c.guardar()
    assert app.configuracion.columnas[i].quien_edita() == "Profesores, Gestores, CASA"
    c.tipo.set(TIPOS["lista"]); c.opciones.set("Solo una"); c.guardar()
    assert app.dialogos["avisos"][-1] == "Columna inválida" and app.configuracion.columnas[i].tipo == "texto"

    c.tabla.selection_set(str(i)); c.mover(-1)
    assert app.configuracion.columnas[i - 1].nombre == "Curso"
    c.tabla.selection_set(str(i - 1)); c.quitar()
    assert "Curso" not in [col.nombre for col in guardada().columnas]


def test_programas(app):
    p = app.programas
    p.nombre.set("NUEVO"); p.agregar()
    assert app.configuracion.programas[-1] == "NUEVO"
    p.nombre.set("Usuarios"); p.agregar()
    assert app.dialogos["avisos"][-1] == "Programa inválido"
    p.lista.selection_set(len(app.configuracion.programas) - 1)
    p.nombre.set("NUEVO-2"); p.renombrar()
    p.lista.selection_set(len(app.configuracion.programas) - 1); p.mover(-1)
    assert app.configuracion.programas[-2] == "NUEVO-2"
    p.lista.selection_set(len(app.configuracion.programas) - 2); p.quitar()
    assert "NUEVO-2" not in guardada().programas


def test_configuracion_invalida_no_genera(app):
    for col in app.configuracion.columnas:
        col.identifica = False
    app.configuracion_cambiada()
    assert "⚠️" in app.estado_config.cget("text")
    app.generar()
    assert app.dialogos["avisos"][-1] == "Configuración incompleta"


def test_exportar_importar_y_generar(app, pequena, tmp_path):
    pequena.guardar_json(tmp_path / "otra.json")
    app.dialogos["abrir"].append(str(tmp_path / "otra.json"))
    app.importar()
    assert app.configuracion == pequena and list(app.programas.lista.get(0, "end")) == ["AAA", "BBB"]
    assert app.opciones.filas.get() == "20"

    app.dialogos["guardar"] += [str(tmp_path / "exportada.json"), str(tmp_path / "salida" / "plantilla.xlsx"),
                                str(tmp_path / "Codigo.gs")]
    app.exportar()
    assert Configuracion.leer_json(tmp_path / "exportada.json") == pequena
    app.generar()
    esperar(app)
    assert "✓ Listo" in app.registro.contenido()
    libro = openpyxl.load_workbook(tmp_path / "salida" / "plantilla.xlsx")
    assert libro.sheetnames == ["Instrucciones", "AAA", "BBB", "Consolidado", "Permisos", "Usuarios"]

    app.opciones.dominio.set("ejemplo.edu.co"); app.opciones.guardar_script()
    app.guardar_script("Codigo.gs")
    assert "DOMINIO: 'ejemplo.edu.co'," in (tmp_path / "Codigo.gs").read_text(encoding="utf-8")

    app.restablecer()
    assert app.configuracion == predeterminada()
