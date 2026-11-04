import pytest

from cronograma_neon.comun.configuracion import Columna, Configuracion


@pytest.fixture(autouse=True)
def ajustes_aislados(tmp_path, monkeypatch):
    """Las preferencias de la interfaz nunca tocan la carpeta real del usuario durante las pruebas."""
    monkeypatch.setenv("CRONOGRAMA_NEON_AJUSTES", str(tmp_path / "ajustes"))


@pytest.fixture
def pequena():
    """Configuración mínima con columnas de cada tipo y permiso."""
    return Configuracion(
        programas=["AAA", "BBB"],
        filas=20,
        columnas=[
            Columna("Curso", profesores=True, gestores=True, identifica=True, ancho=30),
            Columna("Inicio", tipo="fecha", profesores=True, gestores=True),
            Columna("Nivel", tipo="lista", opciones=["Bajo", "Medio", "Alto"], gestores=True),
            Columna("Estudiantes", tipo="numero", profesores=True, gestores=True, casa=True),
            Columna("¿Revisado?", tipo="si_no", casa=True),
            Columna("Nota interna"),  # solo ADMIN
        ],
    )
