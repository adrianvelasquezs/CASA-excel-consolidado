# Cronograma 2027 Bloque Neón — consolidado en Google Sheets

Aplicación de escritorio (Windows y macOS) que genera **desde cero** el libro consolidado para Google Sheets, con una hoja por programa, y entrega el Apps Script que aplica los permisos por cuenta institucional. No necesita archivos de Excel de entrada.

- El equipo CASA define en la app los **programas**, las **columnas** (nombre, tipo, ancho y orden) y **quién puede editar cada columna**: Profesores, Gestores, CASA. ADMIN siempre puede.
- Lo que registra un profesor queda a su nombre y ningún otro profesor puede modificarlo.
- El libro trae una hoja **Permisos** de la que el Apps Script toma los permisos de cada columna. Un ADMIN puede cambiarlos con el libro en uso y aplicarlos con *CASA → Aplicar permisos ahora*.
- La configuración predeterminada reproduce el Cronograma 2027: 12 programas y 34 columnas. Incluye, al final y solo para CASA, **¿Se entregó reporte?** (SI/NO), **Razones de porqué están pendientes reportes** y **Razones de porqué no se realizó medición**.

La guía completa de uso y despliegue, en texto plano, está en [guia-de-despliegue.txt](src/cronograma_neon/guias/guia-de-despliegue.txt). También se abre desde la app (**Ver guía de despliegue**) y se puede descargar con **Guardar guía (.txt)…**.

## Descargar y usar

1. En [Releases](../../releases), descargue el archivo de su sistema:
   - **Windows:** `Cronograma-Neon-<versión>-Windows.exe`. Si aparece "Windows protegió su PC": *Más información → Ejecutar de todas formas*.
   - **macOS:** `Cronograma-Neon-<versión>-macOS.zip`. Descomprímalo y mueva **Cronograma Neon.app** a Aplicaciones. Si macOS la bloquea la primera vez: *Configuración del Sistema → Privacidad y seguridad → Abrir de todas formas*.
2. En la app:
   - **Columnas y permisos:** agregue, quite, edite o reordene columnas y marque quién puede editar cada una.
   - **Programas:** agregue, quite o renombre los programas (una hoja por programa).
   - **Opciones:** filas de registro por programa y opciones del Apps Script.
   - **Generar plantilla para Google Sheets…** y guarde o copie `Codigo.gs` y `appsscript.json`.

**Compartir la configuración:** *Exportar configuración…* genera un `.json` que otra persona carga con *Importar configuración…*. La configuración en edición se guarda sola en la carpeta de preferencias del usuario: `~/Library/Application Support/Cronograma Neon` en macOS, `%APPDATA%\Cronograma Neon` en Windows. Los archivos que se generan se guardan donde usted elija.

## Línea de comandos

El mismo ejecutable (o `uv run cronograma-neon` en desarrollo) acepta comandos:

```bash
uv run cronograma-neon exportar-configuracion --salida configuracion.json
```

```bash
uv run cronograma-neon google --config configuracion.json --salida "Google Sheets.xlsx" --apps-script apps-script
```

El primero guarda la configuración predeterminada para editarla. El segundo genera la plantilla y el Apps Script; sin `--config`, usa la configuración predeterminada. `cronograma-neon verificar` hace una prueba rápida de la instalación y la usa el CI. Use `--help` para ver todas las opciones.

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

El primero abre la interfaz, el segundo corre las pruebas y el tercero compila el ejecutable de su sistema en `dist/`. Las pruebas cubren el modelo de configuración, el generador, la CLI, la interfaz (con los diálogos simulados) y una simulación del Apps Script que requiere Node.js; si Node no está instalado, esa prueba se omite.

```
src/cronograma_neon/
  comun/configuracion.py     Modelo: programas, columnas, permisos y validaciones
  comun/predeterminada.json  Configuración predeterminada (Cronograma 2027)
  comun/estructura.py        Escritura de hojas de programa, Consolidado y Permisos
  comun/ajustes.py           Preferencias del usuario
  google_sheets/             Generador de la plantilla + apps_script/ (Codigo.gs, appsscript.json)
  interfaz/                  Interfaz gráfica (Tkinter)
  guias/                     Guía de despliegue en .txt (se muestra y se descarga desde la app)
  cli.py, __main__.py        Línea de comandos y punto de entrada
tests/                       Pruebas (pytest) y simulación del Apps Script (Node)
empaquetado/                 PyInstaller: spec y lanzador
raw/                         Plantillas originales de 2027 (solo referencia histórica; la app ya no las usa)
.github/workflows/           CI/CD
```

## CI/CD y releases

[`.github/workflows/compilar-y-publicar.yml`](.github/workflows/compilar-y-publicar.yml) corre en cada push a `main`, en cada pull request y a mano (*Run workflow*). Hace lo siguiente en Windows y macOS:

1. Instala las dependencias con uv (`uv.lock`) y corre las pruebas.
2. Compila con PyInstaller: un `.exe` de un solo archivo en Windows y un `.app` en macOS.
3. Ejecuta `verificar` sobre el ejecutable compilado.
4. Sube los ejecutables como artefactos del workflow: el `.exe`, y el `.app` comprimido en `.zip`.

**Publicar una versión:**
1. Cambie `__version__` en `src/cronograma_neon/__init__.py` y haga commit en `main`.
2. Cree y suba el tag con la misma versión (por ejemplo `v3.0.0`):
   ```bash
   git tag v3.0.0
   ```
   ```bash
   git push origin v3.0.0
   ```

El workflow verifica que el tag coincida con la versión y crea el release con los dos ejecutables.

Los ejecutables no están firmados con certificados de Microsoft ni de Apple, y por eso el sistema muestra advertencias la primera vez. Para quitarlas hay que firmar la app de Windows y firmar y notarizar la de macOS, con cuentas de desarrollador de pago.
