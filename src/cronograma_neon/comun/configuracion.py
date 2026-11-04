"""
Configuración del consolidado: programas, columnas y quién puede editar cada columna.

Es la única fuente para generar el libro: no se leen plantillas de Excel. Se
edita desde la interfaz, se guarda en las preferencias del usuario y se puede
importar/exportar como JSON para compartirla.
"""

import copy
import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

PREDETERMINADA = Path(__file__).resolve().parent / "predeterminada.json"
VERSION_FORMATO = 1

FILAS_MIN, FILAS_MAX = 10, 2000
ANCHO_MIN, ANCHO_MAX = 5, 100

TIPOS = {
    "texto": "Texto libre",
    "numero": "Número entero (≥ 0)",
    "fecha": "Fecha (dd/mm/aaaa)",
    "si_no": "SI / NO",
    "lista": "Lista de opciones",
}

# Columnas que agrega el sistema: «Programa» al inicio, propiedad de la fila al final.
COLUMNA_PROGRAMA = "Programa"
COLUMNAS_SISTEMA_FIN = ("Registrado por", "Fecha de registro")
COLUMNAS_SISTEMA = (COLUMNA_PROGRAMA, *COLUMNAS_SISTEMA_FIN)

# Hojas que crea el libro; ningún programa puede llamarse igual.
HOJAS_FIJAS = ("Instrucciones", "Consolidado", "Permisos", "Usuarios", "_Auditoria")


class ErrorGeneracion(Exception):
    """Error con un mensaje pensado para mostrarse tal cual al usuario."""


@dataclass
class Columna:
    nombre: str
    tipo: str = "texto"
    opciones: list = field(default_factory=list)
    ancho: int = 20
    profesores: bool = False
    gestores: bool = False
    casa: bool = False
    identifica: bool = False  # la fila cuenta como «curso registrado» cuando se llena alguna de estas

    @property
    def es_del_programa(self):
        return self.profesores or self.gestores

    def quien_edita(self):
        """Texto corto para mostrar: «Profesores, Gestores», «CASA», «Solo ADMIN»…"""
        partes = [n for n, v in (("Profesores", self.profesores), ("Gestores", self.gestores), ("CASA", self.casa)) if v]
        return ", ".join(partes) if partes else "Solo ADMIN"

    def color(self):
        """Color del encabezado según quién edita."""
        if self.es_del_programa and self.casa:
            return "mixto"
        if self.es_del_programa:
            return "programa"
        if self.casa:
            return "casa"
        return "admin"


@dataclass
class Configuracion:
    programas: list
    columnas: list
    filas: int = 150

    # --- lectura y escritura
    @classmethod
    def desde_dict(cls, datos):
        try:
            columnas = [Columna(
                nombre=str(c["nombre"]),
                tipo=str(c.get("tipo", "texto")),
                opciones=[str(o) for o in c.get("opciones", [])],
                ancho=int(c.get("ancho", 20)),
                profesores=bool(c.get("profesores")),
                gestores=bool(c.get("gestores")),
                casa=bool(c.get("casa")),
                identifica=bool(c.get("identifica")),
            ) for c in datos["columnas"]]
            return cls(programas=[str(p) for p in datos["programas"]], columnas=columnas,
                       filas=int(datos.get("filas", 150)))
        except (KeyError, TypeError, ValueError) as error:
            raise ErrorGeneracion(f"La configuración no tiene el formato esperado ({error}).") from error

    def a_dict(self):
        return {"version": VERSION_FORMATO, "filas": self.filas, "programas": list(self.programas),
                "columnas": [asdict(c) for c in self.columnas]}

    @classmethod
    def leer_json(cls, ruta):
        try:
            datos = json.loads(Path(ruta).read_text(encoding="utf-8"))
        except OSError as error:
            raise ErrorGeneracion(f"No se pudo leer {ruta}: {error}") from error
        except ValueError as error:
            raise ErrorGeneracion(f"{Path(ruta).name} no es un JSON válido: {error}") from error
        configuracion = cls.desde_dict(datos)
        configuracion.validar()
        return configuracion

    def guardar_json(self, ruta):
        try:
            Path(ruta).parent.mkdir(parents=True, exist_ok=True)
            Path(ruta).write_text(json.dumps(self.a_dict(), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        except OSError as error:
            raise ErrorGeneracion(f"No se pudo guardar {ruta}: {error}") from error

    def copia(self):
        return copy.deepcopy(self)

    # --- validación
    def validar(self):
        """Lanza ErrorGeneracion con el primer problema encontrado."""
        validar_filas(self.filas)
        if not self.programas:
            raise ErrorGeneracion("Agregue al menos un programa.")
        vistos = set()
        for programa in self.programas:
            validar_programa(programa)
            if programa.lower() in vistos:
                raise ErrorGeneracion(f"El programa «{programa}» está repetido.")
            vistos.add(programa.lower())

        if not self.columnas:
            raise ErrorGeneracion("Agregue al menos una columna.")
        vistos = set()
        for col in self.columnas:
            validar_columna(col)
            if col.nombre.lower() in vistos:
                raise ErrorGeneracion(f"La columna «{col.nombre}» está repetida.")
            vistos.add(col.nombre.lower())
        if not any(c.identifica for c in self.columnas):
            raise ErrorGeneracion("Marque al menos una columna como «Identifica el curso».")
        return self


def predeterminada():
    """Configuración inicial: los 12 programas y las columnas originales del Cronograma 2027."""
    return Configuracion.leer_json(PREDETERMINADA)


def validar_filas(filas):
    if not isinstance(filas, int) or not FILAS_MIN <= filas <= FILAS_MAX:
        raise ErrorGeneracion(f"Las filas por programa deben estar entre {FILAS_MIN} y {FILAS_MAX}.")


def validar_programa(programa):
    if not programa or programa != programa.strip():
        raise ErrorGeneracion(f"«{programa}» no es un nombre de programa válido (vacío o con espacios al inicio o final).")
    if len(programa) > 31 or re.search(r"[\[\]:*?/\\']", programa):
        raise ErrorGeneracion(f"«{programa}» no sirve como nombre de hoja: máximo 31 caracteres y sin [ ] : * ? / \\ '")
    if programa.lower() in {h.lower() for h in HOJAS_FIJAS} or programa.startswith("_"):
        raise ErrorGeneracion(f"«{programa}» es un nombre reservado del libro.")


def validar_columna(col):
    nombre = col.nombre
    if not nombre or nombre != nombre.strip() or "\n" in nombre:
        raise ErrorGeneracion(f"«{nombre}» no es un nombre de columna válido (vacío, con saltos de línea o espacios al borde).")
    if len(nombre) > 120:
        raise ErrorGeneracion(f"El nombre de columna «{nombre[:40]}…» es demasiado largo (máximo 120 caracteres).")
    if nombre.lower() in {s.lower() for s in COLUMNAS_SISTEMA}:
        raise ErrorGeneracion(f"«{nombre}» es una columna del sistema; use otro nombre.")
    if col.tipo not in TIPOS:
        raise ErrorGeneracion(f"La columna «{nombre}» tiene un tipo desconocido: {col.tipo}.")
    if col.tipo == "lista":
        opciones = [o for o in col.opciones if o.strip()]
        if len(opciones) < 2:
            raise ErrorGeneracion(f"La lista de «{nombre}» necesita al menos dos opciones.")
        if any("," in o for o in opciones):
            raise ErrorGeneracion(f"Las opciones de «{nombre}» no pueden contener comas.")
        if len(",".join(opciones)) > 250:
            raise ErrorGeneracion(f"Las opciones de «{nombre}» son demasiado largas (máximo 250 caracteres en total).")
    if not ANCHO_MIN <= int(col.ancho) <= ANCHO_MAX:
        raise ErrorGeneracion(f"El ancho de «{nombre}» debe estar entre {ANCHO_MIN} y {ANCHO_MAX}.")


def opciones_de(col):
    """Opciones de la lista desplegable de una columna, o None si no tiene lista."""
    if col.tipo == "si_no":
        return ["SI", "NO"]
    if col.tipo == "lista":
        return [o.strip() for o in col.opciones if o.strip()]
    return None
