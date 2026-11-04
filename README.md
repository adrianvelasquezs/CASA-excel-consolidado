# Cronograma 2027 Bloque Neón — consolidado en Google Sheets

Aplicación de escritorio (Windows y macOS) que une las plantillas de cada programa en una sola plantilla para Google Sheets, y entrega el Apps Script que aplica los permisos por cuenta institucional:

- Cada programa edita solo su hoja; los profesores y gestores solo las columnas **azules** y el equipo CASA solo las **verdes**.
- Lo que registra un profesor queda a su nombre y ningún otro profesor puede modificarlo.
- Títulos, columnas del sistema y las hojas Consolidado, Usuarios e Instrucciones son solo para administración.
- A todos los programas se agregan, al final y en verde (solo CASA): **¿Se entregó reporte?** (SI/NO), **Razones de porqué están pendientes reportes** y **Razones de porqué no se realizó medición**.

La guía completa de despliegue está en [google-sheets.md](src/cronograma_neon/guias/google-sheets.md) y también se abre desde la app (**Ver guía de despliegue**).

## Descargar y usar

1. En [Releases](../../releases), descargue el archivo de su sistema:
   - **Windows:** `Cronograma-Neon-<versión>-Windows.exe`. Si aparece "Windows protegió su PC": *Más información → Ejecutar de todas formas*.
   - **macOS:** `Cronograma-Neon-<versión>-macOS.zip`. Descomprímalo y mueva **Cronograma Neon.app** a Aplicaciones. Si macOS la bloquea la primera vez: *Configuración del Sistema → Privacidad y seguridad → Abrir de todas formas*.
2. Abra la app y siga sus secciones:
   1. **Plantillas de los programas:** archivos sueltos o una carpeta. Por defecto se usan las plantillas incluidas en la app.
   2. **Configuración de la plantilla:** filas de registro por programa.
   3. **Configuración del Apps Script:** dominio, hora de restauración y opciones de permisos.
   4. **Generar y descargar:** la plantilla `.xlsx`, `Codigo.gs` y `appsscript.json`.

**No hay rutas fijas.** Cada archivo se carga o se guarda donde usted elija:

| Elemento | Cargar | Guardar |
|---|---|---|
| Plantillas de los programas | *Elegir archivos… / Elegir carpeta…* | — |
| Plantilla para Google Sheets | — | *Generar plantilla para Google Sheets…* |
| `Codigo.gs` / `appsscript.json` | — | *Guardar…* o *Copiar* |

La app recuerda sus preferencias (filas, opciones del Apps Script, plantillas elegidas y última carpeta) en la carpeta de configuración del usuario: `~/Library/Application Support/Cronograma Neon` en macOS, `%APPDATA%\Cronograma Neon` en Windows.

## Línea de comandos

El mismo ejecutable (o `uv run cronograma-neon` en desarrollo) acepta comandos; todas las rutas se indican explícitamente:

```bash
uv run cronograma-neon google --plantillas raw --salida "Google Sheets.xlsx" --apps-script apps-script --dominio uniandes.edu.co
```

`cronograma-neon verificar` hace una prueba rápida de la instalación; la usa el CI. Use `--help` para ver todas las opciones.

## Desarrollo

Requiere [uv](https://docs.astral.sh/uv/) (usa Python 3.12, fijado en `.python-version`).

```bash
uv run cronograma-neon
```

```bash
uv run pytest
```

```bash
uv run --group build pyinstaller --noconfirm --clean empaquetado/cronograma_neon.spec
```

El primero abre la interfaz, el segundo corre las pruebas y el tercero compila el ejecutable de su sistema en `dist/`. Las pruebas cubren el generador, la CLI, la interfaz (con los diálogos simulados) y una simulación del Apps Script que requiere Node.js; si Node no está instalado, esa prueba se omite.

```
raw/                         Plantillas de ejemplo (se incluyen en la app; no se modifican)
src/cronograma_neon/
  comun/                     Lectura de plantillas, columnas adicionales, estructura de las hojas, preferencias
  google_sheets/             Generador de la plantilla + apps_script/ (Codigo.gs, appsscript.json)
  interfaz/                  Interfaz gráfica (Tkinter)
  guias/                     Guía de despliegue (se muestra dentro de la app)
  cli.py, __main__.py        Línea de comandos y punto de entrada
tests/                       Pruebas (pytest) y simulación del Apps Script (Node)
empaquetado/                 PyInstaller: spec y lanzador
.github/workflows/           CI/CD
```

Las columnas que se agregan a todos los programas se definen en `COLUMNAS_ADICIONALES` (`src/cronograma_neon/comun/plantillas.py`). Ahí se indica el nombre, el color (`False` = verde, CASA) y el ancho de cada una.

## CI/CD y releases

[`.github/workflows/compilar-y-publicar.yml`](.github/workflows/compilar-y-publicar.yml) corre en cada push a `main`, en cada pull request y a mano (*Run workflow*). Hace lo siguiente en Windows y macOS:

1. Instala las dependencias con uv (`uv.lock`) y corre las pruebas.
2. Compila con PyInstaller: un `.exe` de un solo archivo en Windows y un `.app` en macOS.
3. Ejecuta `verificar` sobre el ejecutable compilado.
4. Sube los ejecutables como artefactos del workflow: el `.exe`, y el `.app` comprimido en `.zip`.

**Publicar una versión:**
1. Cambie `__version__` en `src/cronograma_neon/__init__.py` y haga commit.
2. Cree y suba el tag con la misma versión:
   ```bash
   git tag v2.0.0
   ```
   ```bash
   git push origin v2.0.0
   ```

El workflow verifica que el tag coincida con la versión y crea el release con los dos ejecutables.

Los ejecutables no están firmados con certificados de Microsoft ni de Apple, y por eso el sistema muestra advertencias la primera vez. Para quitarlas hay que firmar la app de Windows y firmar y notarizar la de macOS, con cuentas de desarrollador de pago.
