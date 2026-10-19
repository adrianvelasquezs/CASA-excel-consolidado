"""Maneja la interfaz por código, con los diálogos de archivos y mensajes simulados."""

import time
import tkinter as tk

import pytest

from cronograma_neon.comun import ajustes
from cronograma_neon.interfaz import app as interfaz
from cronograma_neon.onedrive import generador as onedrive


@pytest.fixture
def app(monkeypatch):
    try:
        ventana = interfaz.App()
    except tk.TclError as error:  # sin pantalla (p. ej. Linux sin X)
        pytest.skip(f"Tk no disponible: {error}")
    ventana.withdraw()
    dialogos = {"guardar": [], "abrir": [], "avisos": [], "respuestas": []}

    def responder(*_a, **_k):
        return dialogos["respuestas"].pop(0) if dialogos["respuestas"] else True

    monkeypatch.setattr(interfaz.filedialog, "asksaveasfilename", lambda **k: dialogos["guardar"].pop(0))
    monkeypatch.setattr(interfaz.filedialog, "askopenfilename", lambda **k: dialogos["abrir"].pop(0))
    for nombre in ("showwarning", "showinfo", "showerror"):
        monkeypatch.setattr(interfaz.messagebox, nombre, lambda titulo, *_a, **_k: dialogos["avisos"].append(titulo))
    monkeypatch.setattr(interfaz.messagebox, "askyesno", responder)
    monkeypatch.setattr(interfaz.messagebox, "askyesnocancel", responder)
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


def test_inicia_con_plantillas_incluidas(app):
    assert len(app.programas) == 12
    assert "incluidas" in app.estado_plantillas.cget("text")


def test_onedrive_generar_y_guardar_contrasenas(app, tmp_path):
    od = app.onedrive
    od.prog.set("AFIN"); od.prof.set("Ana"); od.cant.set("4"); od.agregar_profesor()
    od.prog.set("AFIN"); od.prof.set("ana"); od.cant.set("1"); od.agregar_profesor()
    assert app.dialogos["avisos"][-1] == "Repetido"
    assert ajustes.leer()["onedrive"]["profesores"] == [["AFIN", "Ana", 4]]  # preferencias guardadas fuera del proyecto

    excel, claves = tmp_path / "salida" / "libro.xlsx", tmp_path / "privado" / "claves.csv"
    claves.parent.mkdir()
    app.dialogos["guardar"] += [str(excel), str(claves)]
    od.generar()
    esperar(app)
    assert excel.exists() and "✓ Listo" in od.registro.contenido()
    assert claves.exists() and not od.sin_guardar  # aceptó guardar las contraseñas nuevas
    assert onedrive.leer_contrasenas_csv(claves) == od.contrasenas
    assert len(od.tabla_pw.get_children()) == 15

    # Una sesión nueva carga las contraseñas guardadas y las reutiliza.
    od.contrasenas, od.sin_guardar = {}, False
    app.dialogos["abrir"].append(str(claves))
    od.cargar_contrasenas()
    anteriores = dict(od.contrasenas)
    app.dialogos["guardar"].append(str(excel))
    od.generar()
    esperar(app)
    assert od.contrasenas == anteriores and not od.sin_guardar


def test_regenerar_marca_sin_guardar_y_pide_al_cerrar(app, tmp_path):
    od = app.onedrive
    app.dialogos["guardar"] += [str(tmp_path / "a.xlsx")]
    app.dialogos["respuestas"] += [True, False]  # crear nuevas: sí · guardar contraseñas ahora: no
    od.generar()
    esperar(app)
    assert od.sin_guardar and "SIN GUARDAR" in od.estado_pw.cget("text")

    app.dialogos["respuestas"].append(None)  # Cancelar al cerrar
    app.cerrar()
    assert app.winfo_exists()

    od.tabla_pw.selection_set("programa\x1fMF")
    antes = od.contrasenas[("programa", "MF")]
    app.dialogos["guardar"] += [str(tmp_path / "b.xlsx"), str(tmp_path / "c.csv")]
    od.regenerar(todas=False)
    esperar(app)
    assert od.contrasenas[("programa", "MF")] != antes
    assert onedrive.leer_contrasenas_csv(tmp_path / "c.csv")[("programa", "MF")] == od.contrasenas[("programa", "MF")]


def test_google_generar_y_guardar_script(app, tmp_path):
    gs = app.google
    gs.dominio.set("no valido")
    assert gs.config_valida() is None and "Dominio" in gs.estado_script.cget("text")
    gs.dominio.set("ejemplo.edu.co"); gs.hora.set("3"); gs.guardar_preferencias()
    assert ajustes.leer()["google"]["apps_script"]["DOMINIO"] == "ejemplo.edu.co"

    app.dialogos["guardar"] += [str(tmp_path / "Codigo.gs"), str(tmp_path / "plantilla.xlsx")]
    gs.guardar_script("Codigo.gs")
    assert "DOMINIO: 'ejemplo.edu.co'," in (tmp_path / "Codigo.gs").read_text(encoding="utf-8")
    gs.generar()
    esperar(app)
    assert (tmp_path / "plantilla.xlsx").exists()
    assert "OneDrive" not in gs.registro.contenido() and "Contraseñas" not in gs.registro.contenido()


def test_plantillas_elegidas_por_carpeta(app, plantillas, tmp_path, monkeypatch):
    monkeypatch.setattr(interfaz.filedialog, "askdirectory", lambda **k: str(tmp_path))
    app.elegir_carpeta()  # carpeta vacía
    assert app.programas == [] and "⚠️" in app.estado_plantillas.cget("text")
    app.onedrive.generar()
    assert app.dialogos["avisos"][-1] == "Plantillas"
    monkeypatch.setattr(interfaz.filedialog, "askdirectory", lambda **k: str(plantillas))
    app.elegir_carpeta()
    assert len(app.programas) == 12 and ajustes.leer()["plantillas"] == [str(plantillas)]
    app.usar_incluidas()
    assert ajustes.leer()["plantillas"] is None
