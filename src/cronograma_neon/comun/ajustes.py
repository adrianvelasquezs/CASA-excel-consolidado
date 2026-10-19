"""
Preferencias de la interfaz (filas, profesores, opciones del Apps Script,
plantillas elegidas y última carpeta usada).

Se guardan en la carpeta de configuración del usuario de cada sistema
operativo, nunca dentro del proyecto ni junto al ejecutable. Las contraseñas
NO se guardan aquí: se cargan y guardan explícitamente como archivo.
"""

import copy
import json
import os
import sys
from pathlib import Path

POR_DEFECTO = {
    "plantillas": None,  # None = usar las plantillas incluidas en la app
    "ultima_carpeta": None,
    "onedrive": {"filas": 150, "profesores": []},
    "google": {"filas": 150, "apps_script": {}},
}


def carpeta_ajustes():
    """Carpeta de configuración del usuario. CRONOGRAMA_NEON_AJUSTES permite cambiarla (p. ej. en pruebas)."""
    if os.environ.get("CRONOGRAMA_NEON_AJUSTES"):
        return Path(os.environ["CRONOGRAMA_NEON_AJUSTES"])
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "Cronograma Neon"


def archivo_ajustes():
    return carpeta_ajustes() / "ajustes.json"


def _mezclar(base, nuevos):
    """Aplica `nuevos` sobre los valores por defecto; un diccionario vacío por defecto acepta cualquier clave."""
    if not base:
        base.update(nuevos)
        return base
    for clave, valor in nuevos.items():
        if isinstance(valor, dict) and isinstance(base.get(clave), dict):
            _mezclar(base[clave], valor)
        elif clave in base:
            base[clave] = valor
    return base


def leer():
    ajustes = copy.deepcopy(POR_DEFECTO)
    try:
        return _mezclar(ajustes, json.loads(archivo_ajustes().read_text(encoding="utf-8")))
    except (OSError, ValueError, AttributeError):
        return ajustes


def guardar(ajustes):
    """Guarda sin interrumpir la app si la carpeta no se puede escribir."""
    try:
        archivo_ajustes().parent.mkdir(parents=True, exist_ok=True)
        archivo_ajustes().write_text(json.dumps(ajustes, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        return True
    except OSError:
        return False
