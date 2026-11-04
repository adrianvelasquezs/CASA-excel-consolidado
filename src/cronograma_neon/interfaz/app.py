"""
Interfaz gráfica. No usa rutas fijas: las plantillas se eligen y cada resultado
(plantilla para Google Sheets, Codigo.gs, appsscript.json) se guarda donde el
usuario elija con los diálogos del sistema.
"""

import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from .. import __version__
from ..comun import ajustes as preferencias
from ..comun.plantillas import (
    FILAS_MAX,
    FILAS_MIN,
    ErrorGeneracion,
    carpeta_plantillas_incluidas,
    expandir_rutas,
    leer_plantillas,
    programa_de,
)
from ..google_sheets import generador as google

COLOR = "#188038"
COLOR_ERROR = "#C5221F"
GUIA = Path(__file__).resolve().parent.parent / "guias" / "google-sheets.md"
TIPOS_EXCEL = [("Libro de Excel", "*.xlsx")]


# ---------------------------------------------------------------------------
# Piezas visuales reutilizables
# ---------------------------------------------------------------------------

class BotonColor(tk.Label):
    """Botón grande de color (los botones nativos de Mac no aceptan color de fondo)."""

    def __init__(self, padre, texto, color, comando):
        super().__init__(padre, text=texto, bg=color, fg="white", padx=18, pady=10, cursor="hand2",
                         font=tkfont.Font(weight="bold", size=13))
        self.color, self.comando, self.activo = color, comando, True
        self.bind("<Button-1>", lambda _e: self.activo and self.comando())

    def habilitar(self, activo):
        self.activo = activo
        self.configure(bg=self.color if activo else "#9AA0A6", cursor="hand2" if activo else "arrow")


def nota(padre, texto=""):
    return ttk.Label(padre, text=texto, foreground="#5F6368", wraplength=880, justify="left")


class Registro(ttk.Frame):
    """Caja de texto de solo lectura con el resultado de cada acción."""

    def __init__(self, padre, titulo):
        super().__init__(padre)
        ttk.Label(self, text=titulo).pack(anchor="w")
        self.texto = tk.Text(self, height=5, wrap="word", state="disabled", relief="solid", borderwidth=1)
        self.texto.pack(fill="both", expand=True)

    def escribir(self, mensaje, limpiar=False):
        self.texto.configure(state="normal")
        if limpiar:
            self.texto.delete("1.0", "end")
        self.texto.insert("end", str(mensaje).rstrip() + "\n")
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def contenido(self):
        return self.texto.get("1.0", "end")


def mostrar_texto(padre, titulo, texto):
    """Ventana de solo lectura (guía, lista de plantillas)."""
    ventana = tk.Toplevel(padre)
    ventana.title(titulo)
    ventana.geometry("820x640")
    caja = tk.Text(ventana, wrap="word", padx=12, pady=10)
    barra = ttk.Scrollbar(ventana, orient="vertical", command=caja.yview)
    caja.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    caja.pack(fill="both", expand=True)
    caja.insert("1.0", texto)
    caja.configure(state="disabled")
    return ventana


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Cronograma 2027 Bloque Neón — Google Sheets {__version__}")
        self.ocupado = False
        self.ajustes = preferencias.leer()
        self.programas = []
        conf = self.ajustes["google"]
        try:
            valores = google.validar_config_apps_script(conf["apps_script"])
        except ErrorGeneracion:
            valores = google.config_apps_script_por_defecto()

        marco = tk.Frame(self, bg=COLOR, padx=14, pady=8)
        marco.pack(fill="x")
        tk.Label(marco, text="Cronograma consolidado para Google Sheets", bg=COLOR, fg="white",
                 font=tkfont.Font(size=17, weight="bold")).pack(anchor="w")
        tk.Label(marco, text="Genera la plantilla a importar en Google Sheets y el Apps Script que aplica los permisos.",
                 bg=COLOR, fg="white").pack(anchor="w")
        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)

        # 1. Plantillas de entrada
        entrada = ttk.LabelFrame(cuerpo, text=" 1. Plantillas de los programas ", padding=8)
        entrada.pack(fill="x")
        self.estado_plantillas = ttk.Label(entrada, text="")
        self.estado_plantillas.pack(side="left")
        ttk.Button(entrada, text="Ver lista", command=self.ver_plantillas).pack(side="right")
        ttk.Button(entrada, text="Usar las incluidas", command=self.usar_incluidas).pack(side="right", padx=6)
        ttk.Button(entrada, text="Elegir carpeta…", command=self.elegir_carpeta).pack(side="right")
        ttk.Button(entrada, text="Elegir archivos…", command=self.elegir_archivos).pack(side="right", padx=6)

        # 2. Plantilla
        plantilla = ttk.LabelFrame(cuerpo, text=" 2. Configuración de la plantilla ", padding=8)
        plantilla.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(plantilla)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(conf["filas"]))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_preferencias())
        nota(plantilla, "Los usuarios y sus roles no se configuran aquí: se administran en la hoja «Usuarios» "
                        "dentro de Google Sheets.").pack(anchor="w", pady=(6, 0))

        # 3. Apps Script
        script = ttk.LabelFrame(cuerpo, text=" 3. Configuración del Apps Script ", padding=8)
        script.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(script)
        fila.pack(fill="x")
        ttk.Label(fila, text="Dominio institucional:  @").pack(side="left")
        self.dominio = tk.StringVar(value=valores["DOMINIO"])
        caja_dominio = ttk.Entry(fila, textvariable=self.dominio, width=24)
        caja_dominio.pack(side="left")
        caja_dominio.bind("<FocusOut>", lambda _e: self.guardar_preferencias())
        caja_dominio.bind("<Return>", lambda _e: self.guardar_preferencias())
        ttk.Label(fila, text="      Restaurar permisos cada día a las").pack(side="left")
        self.hora = tk.StringVar(value=str(valores["HORA_RESTAURACION"]))
        ttk.Spinbox(fila, from_=0, to=23, width=4, textvariable=self.hora, wrap=True,
                    command=self.guardar_preferencias).pack(side="left", padx=4)
        ttk.Label(fila, text="h").pack(side="left")
        self.gestores = tk.BooleanVar(value=valores["GESTOR_EDITA_FILAS_DE_PROFESORES"])
        self.alertas = tk.BooleanVar(value=valores["ALERTAS_POR_CORREO"])
        ttk.Checkbutton(script, text="Los gestores pueden editar las filas registradas por sus profesores",
                        variable=self.gestores, command=self.guardar_preferencias).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(script, text="Avisar por correo a los ADMIN cuando alguien edita algo que no le corresponde",
                        variable=self.alertas, command=self.guardar_preferencias).pack(anchor="w")
        self.estado_script = nota(script)
        self.estado_script.pack(anchor="w", pady=(4, 0))

        # 4. Generar
        gen = ttk.LabelFrame(cuerpo, text=" 4. Generar y descargar ", padding=8)
        gen.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(gen)
        fila.pack(fill="x")
        self.boton = BotonColor(fila, "Generar plantilla para Google Sheets…", COLOR, self.generar)
        self.boton.pack(side="left")
        ttk.Button(fila, text="Ver guía de despliegue", command=self.ver_guia).pack(side="left", padx=10)
        fila = ttk.Frame(gen)
        fila.pack(fill="x", pady=(8, 0))
        ttk.Label(fila, text="Apps Script:").pack(side="left")
        ttk.Button(fila, text="Guardar Codigo.gs…", command=lambda: self.guardar_script("Codigo.gs")).pack(side="left", padx=(6, 0))
        ttk.Button(fila, text="Guardar appsscript.json…", command=lambda: self.guardar_script("appsscript.json")).pack(side="left", padx=6)
        ttk.Button(fila, text="Copiar Codigo.gs", command=lambda: self.copiar_script("Codigo.gs")).pack(side="left")
        ttk.Button(fila, text="Copiar appsscript.json", command=lambda: self.copiar_script("appsscript.json")).pack(side="left", padx=6)
        nota(gen, "Importe la plantilla a Google Sheets y pegue el código en Extensiones → Apps Script "
                  "(pasos completos en la guía). Si cambia la configuración, vuelva a copiar Codigo.gs.").pack(anchor="w", pady=(6, 0))

        self.registro = Registro(cuerpo, "Resultado:")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))
        self.cargar_plantillas()

        self.update_idletasks()
        alto = min(self.winfo_reqheight(), self.winfo_screenheight() - 80)
        self.geometry(f"{max(self.winfo_reqwidth(), 900)}x{alto}")
        self.minsize(860, min(alto, 600))

    # --- plantillas de entrada
    def rutas_plantillas(self):
        elegidas = self.ajustes["plantillas"]
        if elegidas:
            return [Path(p) for p in elegidas]
        incluidas = carpeta_plantillas_incluidas()
        return [incluidas] if incluidas else []

    def cargar_plantillas(self):
        """Lee las plantillas elegidas para validarlas y obtener los programas."""
        rutas = self.rutas_plantillas()
        origen = "elegidas por usted" if self.ajustes["plantillas"] else "incluidas en la app"
        try:
            self.programas = list(leer_plantillas(rutas)[0]) if rutas else []
            if not self.programas:
                raise ErrorGeneracion("No hay plantillas: use «Elegir archivos…» o «Elegir carpeta…».")
            self.estado_plantillas.configure(text=f"{len(self.programas)} programas ({origen})", foreground="")
        except ErrorGeneracion as error:
            self.programas = []
            self.estado_plantillas.configure(text=f"⚠️ {error}", foreground=COLOR_ERROR)

    def elegir_archivos(self):
        rutas = filedialog.askopenfilenames(parent=self, title="Elegir plantillas", filetypes=TIPOS_EXCEL,
                                            initialdir=self.carpeta_inicial())
        if rutas:
            self.recordar_carpeta(rutas[0])
            self.ajustes["plantillas"] = [str(r) for r in rutas]
            self.guardar_ajustes()
            self.cargar_plantillas()

    def elegir_carpeta(self):
        ruta = filedialog.askdirectory(parent=self, title="Carpeta con las plantillas", initialdir=self.carpeta_inicial())
        if ruta:
            self.ajustes["ultima_carpeta"] = ruta
            self.ajustes["plantillas"] = [ruta]
            self.guardar_ajustes()
            self.cargar_plantillas()

    def usar_incluidas(self):
        self.ajustes["plantillas"] = None
        self.guardar_ajustes()
        self.cargar_plantillas()

    def ver_plantillas(self):
        try:
            lineas = [f"{programa_de(a):<16} {a}" for a in expandir_rutas(self.rutas_plantillas())]
        except ErrorGeneracion as error:
            lineas = [str(error)]
        mostrar_texto(self, "Plantillas de entrada", "\n".join(lineas))

    # --- configuración
    def filas_validas(self):
        try:
            n = int(self.filas.get())
        except ValueError:
            return None
        return n if FILAS_MIN <= n <= FILAS_MAX else None

    def config(self):
        return {"DOMINIO": self.dominio.get(), "HORA_RESTAURACION": self.hora.get(),
                "GESTOR_EDITA_FILAS_DE_PROFESORES": self.gestores.get(), "ALERTAS_POR_CORREO": self.alertas.get()}

    def config_valida(self):
        """Configuración validada, o None (con el error visible) si algo está mal."""
        try:
            valores = google.validar_config_apps_script(self.config())
        except ErrorGeneracion as error:
            self.estado_script.configure(text=f"⚠️ {error}", foreground=COLOR_ERROR)
            return None
        self.estado_script.configure(text="✓ Configuración guardada.", foreground=COLOR)
        return valores

    def guardar_preferencias(self):
        if self.filas_validas():
            self.ajustes["google"]["filas"] = self.filas_validas()
        valores = self.config_valida()
        if valores:
            self.ajustes["google"]["apps_script"] = valores
        self.guardar_ajustes()

    # --- Apps Script
    def texto_script(self, nombre):
        if nombre == "appsscript.json":
            return google.manifiesto_apps_script()
        valores = self.config_valida()
        return google.codigo_apps_script(valores) if valores else None

    def guardar_script(self, nombre):
        texto = self.texto_script(nombre)
        if texto is None:
            return
        tipos = [("Apps Script", "*.gs")] if nombre.endswith(".gs") else [("JSON", "*.json")]
        ruta = self.pedir_guardar(f"Guardar {nombre}", nombre, Path(nombre).suffix, tipos)
        if not ruta:
            return
        try:
            google.guardar_texto(ruta, texto)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo guardar", str(error), parent=self)
            return
        self.registro.escribir(f"{nombre} guardado en {ruta}.")

    def copiar_script(self, nombre):
        texto = self.texto_script(nombre)
        if texto is not None:
            self.copiar(texto)
            self.registro.escribir(f"{nombre} copiado al portapapeles. Péguelo en Extensiones → Apps Script.")

    # --- generar
    def generar(self):
        filas, valores = self.filas_validas(), self.config_valida()
        if not filas:
            messagebox.showwarning("Filas inválidas", f"Las filas deben estar entre {FILAS_MIN} y {FILAS_MAX}.", parent=self)
            return
        if not valores:
            return
        if not self.programas:
            messagebox.showwarning("Plantillas", "Primero elija plantillas válidas (sección 1).", parent=self)
            return
        destino = self.pedir_guardar("Guardar plantilla para Google Sheets", google.NOMBRE_SUGERIDO, ".xlsx", TIPOS_EXCEL)
        if not destino:
            return
        plantillas = self.rutas_plantillas()

        def despues(mensajes):
            for m in mensajes:
                self.registro.escribir(m)
            self.registro.escribir("Siguiente: guarde o copie Codigo.gs y appsscript.json (botones de arriba).")

        self.ejecutar(lambda: google.generar_plantilla(plantillas, destino, filas, valores), despues)

    def ejecutar(self, funcion, al_terminar):
        """Corre `funcion` en segundo plano; el resultado vuelve al hilo de la ventana mediante una cola."""
        if self.ocupado:
            return
        self.ocupado = True
        self.boton.habilitar(False)
        self.registro.escribir("Generando plantilla para Google Sheets…", limpiar=True)
        resultado = queue.Queue()

        def trabajo():
            try:
                resultado.put(("ok", funcion()))
            except ErrorGeneracion as error:
                resultado.put(("error", str(error)))
            except Exception:  # noqa: BLE001 - se muestra el detalle para poder reportarlo
                resultado.put(("fallo", traceback.format_exc()))

        def revisar():
            try:
                estado, valor = resultado.get_nowait()
            except queue.Empty:
                self.after(100, revisar)
                return
            self.ocupado = False
            self.boton.habilitar(True)
            if estado == "ok":
                al_terminar(valor)
                self.registro.escribir("✓ Listo.")
            else:
                self.registro.escribir(f"✗ {valor}")
                messagebox.showerror("No se pudo generar", valor if estado == "error" else
                                     "Ocurrió un error inesperado. El detalle quedó en el cuadro de resultado.", parent=self)

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(100, revisar)

    # --- utilidades
    def carpeta_inicial(self):
        carpeta = self.ajustes.get("ultima_carpeta")
        return carpeta if carpeta and Path(carpeta).is_dir() else str(Path.home())

    def recordar_carpeta(self, ruta):
        self.ajustes["ultima_carpeta"] = str(Path(ruta).parent)
        self.guardar_ajustes()

    def pedir_guardar(self, titulo, nombre, extension, tipos):
        ruta = filedialog.asksaveasfilename(parent=self, title=titulo, initialfile=nombre, defaultextension=extension,
                                            filetypes=tipos, initialdir=self.carpeta_inicial())
        if ruta:
            self.recordar_carpeta(ruta)
        return ruta or None

    def guardar_ajustes(self):
        preferencias.guardar(self.ajustes)

    def copiar(self, texto):
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.update()

    def ver_guia(self):
        try:
            texto = GUIA.read_text(encoding="utf-8")
        except OSError:
            texto = "No se encontró la guía en esta instalación."
        mostrar_texto(self, "Guía de despliegue", texto)


def ejecutar():
    App().mainloop()
