# Cronograma 2027 Bloque Neón — consolidado

Un solo libro donde cada programa registra sus cursos y el equipo CASA completa su parte, con permisos de edición separados. Hay dos versiones independientes que parten de las mismas plantillas de `raw/`:

| | **Excel en OneDrive** → [onedrive/](onedrive/README.md) | **Google Sheets** → [google-sheets/](google-sheets/README.md) |
|---|---|---|
| Cómo se identifica a cada persona | Contraseña por rango | Cuenta institucional @uniandes.edu.co |
| Títulos y columnas de sistema bloqueados | ✅ | ✅ |
| Cada programa solo edita su hoja | ✅ (contraseña del programa) | ✅ (por correo) |
| Profesores solo azul, CASA solo verde | ✅ | ✅ |
| Lo que crea un profesor no lo edita otro | ⚠️ Aproximado: filas reservadas por adelantado, contraseña por profesor | ✅ La fila queda a nombre de quien la crea, automáticamente |
| Si alguien se salta la protección | Se detecta con *Mostrar cambios* (revisión manual) | El script revierte, registra y avisa |
| Se puede compartir la clave | Sí (es una contraseña) | No (es la cuenta de cada persona) |

## Estructura

```
raw/                  Entrada común: plantillas originales por programa (NO se modifican)
comun/
  plantillas.py       Código compartido: lee raw/ y arma la estructura de las hojas

onedrive/             ── Versión Excel / OneDrive ──
  generar_excel.py
  config/             Entradas propias (profesores.csv)
  salida/             Excel listo para subir a OneDrive
  privado/            Contraseñas generadas (ignorado por git)

google-sheets/        ── Versión Google Sheets ──
  generar_plantilla.py
  apps-script/        Código que se pega en Extensiones → Apps Script
  salida/             Plantilla lista para importar a Google Sheets

pyproject.toml, uv.lock   Dependencias (uv)
```

Cada versión lee solo `raw/` y su propia carpeta, y escribe solo en su propia carpeta.

## Generar

Requiere [uv](https://docs.astral.sh/uv/). Desde la raíz del proyecto:

```bash
uv run onedrive/generar_excel.py
```

```bash
uv run google-sheets/generar_plantilla.py
```

Las dos aceptan `--filas N` (filas de registro por programa; por defecto 150). Las instrucciones de despliegue están en el README de cada carpeta.
