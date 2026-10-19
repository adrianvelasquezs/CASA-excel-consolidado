from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent


@pytest.fixture
def plantillas():
    """Carpeta con las plantillas de ejemplo del repositorio."""
    return RAIZ / "raw"


@pytest.fixture(autouse=True)
def ajustes_aislados(tmp_path, monkeypatch):
    """Las preferencias de la interfaz nunca tocan la carpeta real del usuario durante las pruebas."""
    monkeypatch.setenv("CRONOGRAMA_NEON_AJUSTES", str(tmp_path / "ajustes"))
