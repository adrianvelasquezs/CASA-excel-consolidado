# Cronograma 2027 Bloque Neón — consolidado

Aplicación de escritorio (Windows y macOS) que une las plantillas de cada programa en un solo libro, con permisos de edición separados para programas, profesores y el equipo CASA. Genera dos versiones independientes:

| | **Excel para OneDrive** | **Google Sheets** |
|---|---|---|
| Cómo se identifica a cada persona | Contraseña por rango | Cuenta institucional @uniandes.edu.co |
| Títulos y columnas de sistema bloqueados | ✅ | ✅ |
| Cada programa solo edita su hoja | ✅ (contraseña del programa) | ✅ (por correo) |
| Profesores solo azul, CASA solo verde | ✅ | ✅ |
| Lo que crea un profesor no lo edita otro | ⚠️ Aproximado: filas reservadas por adelantado, contraseña por profesor | ✅ La fila queda a nombre de quien la crea, automáticamente |
| Si alguien se salta la protección | Se detecta con *Mostrar cambios* (revisión manual) | El script revierte, registra y avisa |
| Guía de despliegue | [onedrive.md](src/cronograma_neon/guias/onedrive.md) | [google-sheets.md](src/cronograma_neon/guias/google-sheets.md) |

## Descargar y usar

1. En [Releases](../../releases), descargue el archivo de su sistema:
   - **Windows:** `Cronograma-Neon-<versión>-Windows.exe`. Si aparece "Windows protegió su PC": *Más información → Ejecutar de todas formas*.
   - **macOS:** `Cronograma-Neon-<versión>-macOS.zip`. Descomprímalo y mueva **Cronograma Neon.app** a Aplicaciones. Si macOS la bloquea la primera vez: *Configuración del Sistema → Privacidad y seguridad → Abrir de todas formas*.
2. Abra la app.
   - Arriba elija las plantillas de los programas: archivos sueltos o una carpeta. Por defecto se usan las plantillas incluidas en la app.
   - Use la pestaña de la versión que necesita. Las guías completas se abren desde el botón **Ver guía de despliegue** de cada pestaña.

**No hay rutas fijas.** Cada elemento se carga y se guarda donde usted elija:

| Elemento | Pestaña | Cargar | Guardar |
|---|---|---|---|
| Plantillas de los programas | (arriba, común) | *Elegir archivos… / Elegir carpeta…* | — |
| Profesores con filas reservadas | OneDrive | *Importar CSV…* | *Exportar CSV…* |
| Contraseñas | OneDrive | *Cargar archivo…* | *Guardar archivo…* (también se ofrece al generar y al cerrar) |
| Excel protegido | OneDrive | — | *Generar Excel para OneDrive…* |
| Plantilla para Google Sheets | Google Sheets | — | *Generar plantilla para Google Sheets…* |
| `Codigo.gs` / `appsscript.json` | Google Sheets | — | *Guardar…* o *Copiar* |

La app recuerda las preferencias de uso (filas, profesores, opciones del Apps Script, plantillas elegidas y última carpeta) en la carpeta de configuración del usuario: `~/Library/Application Support/Cronograma Neon` en macOS, `%APPDATA%\Cronograma Neon` en Windows. **Las contraseñas nunca se guardan ahí**: solo en el archivo que usted elija.

## Línea de comandos

El mismo ejecutable (o `uv run cronograma-neon` en desarrollo) acepta comandos; todas las rutas se indican explícitamente:

```bash
uv run cronograma-neon onedrive --plantillas raw --salida "Consolidado.xlsx" --contrasenas privado/contrasenas.csv --profesores ejemplos/profesores.ejemplo.csv
```

```bash
uv run cronograma-neon google --plantillas raw --salida "Google Sheets.xlsx" --apps-script apps-script --dominio uniandes.edu.co
```

`--contrasenas` se lee si existe y se actualiza con las nuevas. `cronograma-neon verificar` hace una prueba rápida de la instalación; la usa el CI. Use `--help` en cada comando para ver todas las opciones.

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

El primero abre la interfaz, el segundo corre las pruebas y el tercero compila el ejecutable de su sistema en `dist/`. Las pruebas cubren los generadores, la CLI y la interfaz (con los diálogos simulados), además de la simulación del Apps Script; esta última requiere Node.js y se omite si no está instalado.

```
raw/                         Plantillas de ejemplo (se incluyen en la app; no se modifican)
src/cronograma_neon/
  comun/                     Lectura de plantillas, estructura común, preferencias
  onedrive/                  ── Versión Excel / OneDrive (contraseñas, profesores, libro protegido)
  google_sheets/             ── Versión Google Sheets (plantilla) + apps_script/ (Codigo.gs, appsscript.json)
  interfaz/                  Interfaz gráfica (Tkinter)
  guias/                     Guías de despliegue (se muestran dentro de la app)
  cli.py, __main__.py        Línea de comandos y punto de entrada
tests/                       Pruebas (pytest) y simulación del Apps Script (Node)
empaquetado/                 PyInstaller: spec y lanzador
ejemplos/                    CSV de ejemplo para importar profesores
.github/workflows/           CI/CD
```

## CI/CD y releases

[`.github/workflows/compilar-y-publicar.yml`](.github/workflows/compilar-y-publicar.yml) corre en cada push a `main`, en cada pull request y a mano (*Run workflow*). Hace lo siguiente en Windows y macOS:

1. Instala las dependencias con uv (`uv.lock`) y corre las pruebas.
2. Compila con PyInstaller: un `.exe` de un solo archivo en Windows y un `.app` en macOS.
3. Ejecuta `verificar` sobre el ejecutable compilado.
4. Sube los ejecutables como artefactos del workflow: el `.exe`, y el `.app` comprimido en `.zip`.

**Publicar una versión:**
1. Cambie `__version__` en `src/cronograma_neon/__init__.py` (por ejemplo, `1.1.0`) y haga commit.
2. Cree y suba el tag:
   ```bash
   git tag v1.1.0
   ```
   ```bash
   git push origin v1.1.0
   ```

El workflow verifica que el tag coincida con la versión y crea el release con los dos ejecutables. El repositorio debe permitir que GitHub Actions cree releases: *Settings → Actions → General → Workflow permissions*. El workflow ya pide `contents: write`, pero la organización puede restringirlo.

Los ejecutables no están firmados con certificados de Microsoft ni de Apple, y por eso el sistema muestra advertencias la primera vez. Para quitarlas hay que firmar la app de Windows y firmar y notarizar la de macOS, con cuentas de desarrollador de pago.
