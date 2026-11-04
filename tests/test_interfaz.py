"""Maneja la interfaz por código, con los diálogos de archivos y mensajes simulados."""

import time
import tkinter as tk

import openpyxl
import pytest

from cronograma_neon.comun import ajustes
from cronograma_neon.comun.plantillas import COLUMNAS_ADICIONALES
from cronograma_neon.interfaz import app as interfaz


@pytest.fixture
def app(monkeypatch):
    try:
        ventana = interfaz.App()
    except tk.TclError as error:  # sin pantalla (p. ej. Linux sin X)
        pytest.skip(f"Tk no disponible: {error}")
    ventana.withdraw()
    dialogos = {"guardar": [], "avisos": []}
    monkeypatch.setattr(interfaz.filedialog, "asksaveasfilename", lambda **k: dialogos["guardar"].pop(0))
    for nombre in ("showwarning", "showinfo", "showerror"):
        monkeypatch.setattr(interfaz.messagebox, nombre, lambda titulo, *_a, **_k: dialogos["avisos"].append(titulo))
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
    assert not hasattr(app, "onedrive")


def test_generar_y_guardar_script(app, tmp_path):
    app.dominio.set("no valido")
    assert app.config_valida() is None and "Dominio" in app.estado_script.cget("text")
    app.dominio.set("ejemplo.edu.co"); app.hora.set("3"); app.guardar_preferencias()
    assert ajustes.leer()["google"]["apps_script"]["DOMINIO"] == "ejemplo.edu.co"  # fuera del proyecto

    app.dialogos["guardar"] += [str(tmp_path / "Codigo.gs"), str(tmp_path / "salida" / "plantilla.xlsx")]
    app.guardar_script("Codigo.gs")
    assert "DOMINIO: 'ejemplo.edu.co'," in (tmp_path / "Codigo.gs").read_text(encoding="utf-8")
    app.copiar_script("appsscript.json")
    assert '"timeZone"' in app.clipboard_get()

    app.generar()
    esperar(app)
    plantilla = tmp_path / "salida" / "plantilla.xlsx"
    assert plantilla.exists() and "✓ Listo" in app.registro.contenido()
    encabezados = [c.value for c in openpyxl.load_workbook(plantilla)["AFIN"][3]]
    assert all(n in encabezados for n, _, _ in COLUMNAS_ADICIONALES)


def test_filas_invalidas(app):
    app.filas.set("5")
    app.generar()
    assert app.dialogos["avisos"][-1] == "Filas inválidas"


def test_plantillas_elegidas_por_carpeta(app, plantillas, tmp_path, monkeypatch):
    monkeypatch.setattr(interfaz.filedialog, "askdirectory", lambda **k: str(tmp_path))
    app.elegir_carpeta()  # carpeta vacía
    assert app.programas == [] and "⚠️" in app.estado_plantillas.cget("text")
    app.generar()
    assert app.dialogos["avisos"][-1] == "Plantillas"
    monkeypatch.setattr(interfaz.filedialog, "askdirectory", lambda **k: str(plantillas))
    app.elegir_carpeta()
    assert len(app.programas) == 12 and ajustes.leer()["plantillas"] == [str(plantillas)]
    app.usar_incluidas()
    assert ajustes.leer()["plantillas"] is None
