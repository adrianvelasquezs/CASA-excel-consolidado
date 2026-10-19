# Cronograma 2027 Bloque Neón — consolidado

Un solo libro donde cada programa registra sus cursos y el equipo CASA completa su parte, con permisos de edición separados. Hay dos versiones con la misma estructura:

| | **Excel en OneDrive** (esta guía) | **Google Sheets** ([google-sheets/README.md](google-sheets/README.md)) |
|---|---|---|
| Cómo se identifica a cada persona | Contraseña por rango | Cuenta institucional @uniandes.edu.co |
| Títulos y columnas de sistema bloqueados | ✅ | ✅ |
| Cada programa solo edita su hoja | ✅ (contraseña del programa) | ✅ (por correo) |
| Profesores solo azul, CASA solo verde | ✅ | ✅ |
| Lo que crea un profesor no lo edita otro | ⚠️ Aproximado: filas reservadas por adelantado, contraseña por profesor | ✅ La fila queda a nombre de quien la crea, automáticamente |
| Si alguien se salta la protección | Se detecta con *Mostrar cambios* (revisión manual) | El script revierte, registra y avisa |
| Se puede compartir la clave | Sí (es una contraseña) | No (es la cuenta de cada persona) |

## Estructura del directorio

```
raw/                       Plantillas originales por programa (NO se modifican)
config/
  profesores.csv           Excel: filas reservadas para cada profesor (opcional)
  profesores.ejemplo.csv   Ejemplo de cómo llenarlo
scripts/
  plantillas.py            Lectura de raw/ y estructura común de las hojas
  generar_excel.py         Genera la versión Excel / OneDrive
  generar_google_sheets.py Genera la plantilla para Google Sheets
google-sheets/
  Codigo.gs, appsscript.json  Apps Script de permisos para Google Sheets
  README.md                Guía de despliegue en Google Workspace
salida/
  Cronograma 2027 Bloque Neón - Consolidado.xlsx    ← versión Excel (se sube a OneDrive)
  Cronograma 2027 Bloque Neón - Google Sheets.xlsx  ← plantilla para importar a Google Sheets
privado/
  contrasenas.csv          Contraseñas de la versión Excel. NO subir ni compartir completo
pyproject.toml             Dependencias del proyecto (gestionadas con uv)
uv.lock                    Versiones exactas de las dependencias
```

---

# Versión Excel (OneDrive / SharePoint)

## Qué contiene el Excel

| Hoja | Contenido |
|---|---|
| **Instrucciones** | Explicación para usuarios (solo lectura) |
| **AFIN, EMBA, … (12 hojas)** | Una hoja por programa con las 32 columnas originales y 150 filas de registro |
| **Consolidado** | Une automáticamente todo lo registrado en las 12 hojas (solo lectura, requiere Excel 365 / Excel web) |

Se respeta el código de colores de las plantillas: **azul = Programa**, **verde = CASA**. EMBA no tenía la columna "Nombre actividad"; se agregó para que todas las hojas sean iguales.

## Modelo de permisos

| Zona | Quién puede editar | Cómo |
|---|---|---|
| Títulos, leyenda, encabezados, columna "Programa", hoja Consolidado, Instrucciones | Nadie (solo administración) | Hoja protegida con contraseña `admin` |
| Agregar/borrar/renombrar hojas | Nadie (solo administración) | Estructura del libro protegida con `admin` |
| Columnas **azules** de la hoja de un programa | Solo ese programa | Rango con la contraseña de ese programa |
| Filas asignadas a un profesor (azules) | Solo ese profesor | Rango con contraseña personal; el programa ya no puede editarlas y el nombre del profesor queda fijo |
| Columnas **verdes** de todas las hojas | Solo CASA | Rango con la contraseña `casa` |

Cuando alguien intenta escribir en una celda de color, Excel le pide la contraseña del rango. Con la de AFIN no se puede editar MF, y con la de un profesor no se pueden editar las filas de otro.

### Asignar filas a profesores (requisito: lo que crea un profesor no lo modifica otro)

Excel no sabe qué persona está editando: solo controla con contraseñas. Por eso la propiedad de las filas se logra **reservando filas para cada profesor**, cada una con su propia contraseña:

1. Llenar `config/profesores.csv` (ver `profesores.ejemplo.csv`):
   ```
   programa,profesor,filas
   AFIN,Nombre del profesor,5
   ```
   Las filas se asignan en orden desde la fila 4 de la hoja del programa. Las filas sobrantes quedan para el programa (gestores).
2. Volver a generar el archivo (siguiente sección). Las contraseñas existentes se conservan y solo se crean las nuevas.

Si un programa no tiene profesores en el CSV, toda su zona azul se edita con la contraseña del programa.

## Regenerar el archivo

Requiere [uv](https://docs.astral.sh/uv/). `uv run` crea el entorno e instala las dependencias de `uv.lock` automáticamente la primera vez:

```bash
uv run scripts/generar_excel.py --filas 150
```

(La plantilla de Google Sheets se genera con `uv run scripts/generar_google_sheets.py`.)

⚠️ Regenerar crea un archivo **vacío**. Hágalo antes de empezar a registrar datos. Una vez el archivo esté en uso en OneDrive, los cambios de permisos se hacen a mano en Excel (*Revisar → Permitir modificar rangos* / *Administrar protección*) con la contraseña `admin`.

## Despliegue en OneDrive / SharePoint

1. Subir `salida/Cronograma 2027 Bloque Neón - Consolidado.xlsx` a una carpeta de OneDrive o SharePoint de CASA.
2. Compartir con **permiso de edición** a gestores, profesores y equipo CASA. (Si alguien solo debe consultar, se comparte como "Puede ver".)
3. Entregar a cada persona **solo su contraseña**, por un canal privado (correo o Teams directo), tomándola de `privado/contrasenas.csv`:
   - cada programa recibe la suya, cada profesor la suya,
   - el equipo CASA recibe la de `casa`,
   - la de `admin` solo la guarda la persona administradora.
4. Activar el control de cambios: en Excel web, **Revisar → Mostrar cambios** muestra quién modificó cada celda y cuándo. Junto con el **Historial de versiones** de OneDrive permite ver y revertir cualquier edición indebida.

## Limitaciones (importante)

- **La protección de Excel es un control de uso, no seguridad fuerte.** Evita errores y ediciones en la zona equivocada, pero una persona con conocimientos técnicos y permiso de edición podría saltársela, o compartir su contraseña. El respaldo real es *Mostrar cambios* + *Historial de versiones* (que registran el usuario real de Microsoft 365).
- La propiedad de filas por profesor exige saber de antemano cuántas filas necesita cada uno. Si un profesor necesita más, el administrador amplía su rango en *Administrar protección*.
- La hoja Consolidado usa `VSTACK`/`FILTER` (Excel 365 y Excel para la web). En Excel 2019 o anterior aparece como error.
- Probado en Excel para Mac: la protección de hojas y del libro queda activa y la hoja Consolidado reúne bien los datos de varios programas. **No se probó** en la interfaz que Excel pida la contraseña de cada rango, ni en Excel para la web; conviene hacer una prueba rápida con dos o tres usuarios antes de abrirlo a todos.

## Si se necesita control por usuario real

Si hace falta que el sistema reconozca a cada persona sin usar contraseñas:

- **Microsoft 365 (sin salir de OneDrive):** un archivo por programa, cada uno en una carpeta de SharePoint con permisos solo para ese programa y CASA, y un libro maestro de solo lectura que los une con Power Query. Así cada programa queda aislado por su cuenta institucional. Para filas por profesor, la herramienta indicada es **Microsoft Lists** con el permiso de lista "Crear elementos y editar solo los elementos creados por el usuario".
- **Google Sheets:** ya está implementada en [google-sheets/](google-sheets/README.md). Los permisos van por correo y un Apps Script bloquea cada fila para quien la creó.
