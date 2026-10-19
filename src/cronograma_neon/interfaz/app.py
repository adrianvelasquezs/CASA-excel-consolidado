"""
Interfaz gráfica. No usa rutas fijas: cada entrada se carga y cada resultado se
guarda donde el usuario elija con los diálogos del sistema.

Dos pestañas separadas:
  - Excel para OneDrive: profesores, contraseñas y Excel protegido.
  - Google Sheets: plantilla, opciones del Apps Script y su código.
Las plantillas de entrada (arriba) son comunes a las dos.
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
from ..comun.plantillas import FILAS_MAX, FILAS_MIN, ErrorGeneracion, carpeta_plantillas_incluidas, leer_plantillas
from ..google_sheets import generador as google
from ..onedrive import generador as onedrive

COLOR_ONEDRIVE = "#0F6CBD"
COLOR_GOOGLE = "#188038"
COLOR_ERROR = "#C5221F"
GUIAS = Path(__file__).resolve().parent.parent / "guias"
TIPOS_EXCEL = [("Libro de Excel", "*.xlsx")]
TIPOS_CSV = [("CSV", "*.csv")]


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


def encabezado(padre, color, titulo, subtitulo):
    marco = tk.Frame(padre, bg=color, padx=14, pady=8)
    tk.Label(marco, text=titulo, bg=color, fg="white", font=tkfont.Font(size=17, weight="bold")).pack(anchor="w")
    tk.Label(marco, text=subtitulo, bg=color, fg="white", justify="left").pack(anchor="w")
    return marco


def nota(padre, texto=""):
    return ttk.Label(padre, text=texto, foreground="#5F6368", wraplength=880, justify="left")


def tabla(padre, columnas, alto):
    """Treeview con barra de desplazamiento. columnas: [(id, título, ancho, alineación)]."""
    marco = ttk.Frame(padre)
    arbol = ttk.Treeview(marco, columns=[c[0] for c in columnas], show="headings", height=alto)
    for ident, titulo, ancho, alineacion in columnas:
        arbol.heading(ident, text=titulo)
        arbol.column(ident, width=ancho, anchor=alineacion)
    arbol.pack(side="left", fill="x", expand=True)
    barra = ttk.Scrollbar(marco, orient="vertical", command=arbol.yview)
    barra.pack(side="left", fill="y")
    arbol.configure(yscrollcommand=barra.set)
    return marco, arbol


class Registro(ttk.Frame):
    """Caja de texto de solo lectura con el resultado de cada acción."""

    def __init__(self, padre, titulo):
        super().__init__(padre)
        ttk.Label(self, text=titulo).pack(anchor="w")
        self.texto = tk.Text(self, height=4, wrap="word", state="disabled", relief="solid", borderwidth=1)
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
    """Ventana de solo lectura (guías, lista de plantillas)."""
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
# Pestaña OneDrive
# ---------------------------------------------------------------------------

class PestanaOneDrive(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=(0, 0, 0, 8))
        self.app = app
        self.contrasenas = {}
        self.sin_guardar = False
        self.mostrar = False
        conf_od = app.ajustes["onedrive"]

        encabezado(self, COLOR_ONEDRIVE, "Excel para OneDrive / SharePoint",
                   "Permisos con contraseñas. Genera un .xlsx protegido que se sube a OneDrive.").pack(fill="x")
        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)

        # 1. Configuración
        conf = ttk.LabelFrame(cuerpo, text=" 1. Configuración ", padding=8)
        conf.pack(fill="x")
        fila = ttk.Frame(conf)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(conf_od["filas"]))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_preferencias())

        fila = ttk.Frame(conf)
        fila.pack(fill="x", pady=(8, 2))
        ttk.Label(fila, text="Profesores con filas reservadas (cada uno recibe su propia contraseña):").pack(side="left")
        ttk.Button(fila, text="Exportar CSV…", command=self.exportar_profesores).pack(side="right")
        ttk.Button(fila, text="Importar CSV…", command=self.importar_profesores).pack(side="right", padx=6)
        marco, self.tabla_prof = tabla(conf, [("programa", "Programa", 140, "w"), ("profesor", "Profesor", 420, "w"),
                                              ("filas", "Filas", 70, "center")], 3)
        marco.pack(fill="x")
        form = ttk.Frame(conf)
        form.pack(fill="x", pady=(6, 0))
        self.prog = tk.StringVar()
        self.prof = tk.StringVar()
        self.cant = tk.StringVar(value="5")
        self.combo_prog = ttk.Combobox(form, textvariable=self.prog, values=app.programas, state="readonly", width=14)
        self.combo_prog.pack(side="left")
        ttk.Entry(form, textvariable=self.prof, width=26).pack(side="left", padx=6)
        ttk.Spinbox(form, from_=1, to=FILAS_MAX, width=5, textvariable=self.cant).pack(side="left")
        ttk.Button(form, text="Agregar", command=self.agregar_profesor).pack(side="left", padx=6)
        ttk.Button(form, text="Quitar seleccionado", command=self.quitar_profesor).pack(side="left")
        for programa, profesor, cantidad in conf_od["profesores"]:
            self.tabla_prof.insert("", "end", values=(programa, profesor, cantidad))

        # 2. Contraseñas
        cont = ttk.LabelFrame(cuerpo, text=" 2. Contraseñas ", padding=8)
        cont.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(cont)
        fila.pack(fill="x", pady=(0, 4))
        self.estado_pw = ttk.Label(fila, text="")
        self.estado_pw.pack(side="left")
        ttk.Button(fila, text="Guardar archivo…", command=self.guardar_contrasenas).pack(side="right")
        ttk.Button(fila, text="Cargar archivo…", command=self.cargar_contrasenas).pack(side="right", padx=6)
        marco, self.tabla_pw = tabla(cont, [("rol", "Rol", 140, "w"), ("ambito", "Para", 330, "w"),
                                            ("contrasena", "Contraseña", 160, "w")], 3)
        marco.pack(fill="x")
        botones = ttk.Frame(cont)
        botones.pack(fill="x", pady=(6, 0))
        self.btn_mostrar = ttk.Button(botones, text="Mostrar contraseñas", command=self.alternar_mostrar)
        self.btn_mostrar.pack(side="left")
        ttk.Button(botones, text="Copiar seleccionada", command=self.copiar_contrasena).pack(side="left", padx=6)
        ttk.Button(botones, text="Regenerar seleccionada", command=lambda: self.regenerar(todas=False)).pack(side="left")
        ttk.Button(botones, text="Regenerar todas", command=lambda: self.regenerar(todas=True)).pack(side="left", padx=6)

        # 3. Generar
        gen = ttk.LabelFrame(cuerpo, text=" 3. Generar ", padding=8)
        gen.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(gen)
        fila.pack(fill="x")
        self.boton = BotonColor(fila, "Generar Excel para OneDrive…", COLOR_ONEDRIVE, self.generar)
        self.boton.pack(side="left")
        ttk.Button(fila, text="Ver guía de despliegue", command=lambda: app.ver_guia("onedrive.md", "Guía OneDrive")).pack(side="left", padx=10)
        nota(gen, "Le preguntará dónde guardar el Excel (sale vacío, listo para subir a OneDrive) y, si hay "
                  "contraseñas nuevas, dónde guardarlas.").pack(anchor="w", pady=(6, 0))

        self.registro = Registro(cuerpo, "Resultado (OneDrive):")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))
        self.actualizar_contrasenas()

    # --- preferencias y profesores
    def filas_validas(self):
        try:
            n = int(self.filas.get())
        except ValueError:
            return None
        return n if FILAS_MIN <= n <= FILAS_MAX else None

    def profesores(self):
        return [(p, n, int(c)) for p, n, c in (self.tabla_prof.item(i, "values") for i in self.tabla_prof.get_children())]

    def guardar_preferencias(self):
        if self.filas_validas():
            self.app.ajustes["onedrive"]["filas"] = self.filas_validas()
        self.app.ajustes["onedrive"]["profesores"] = [list(p) for p in self.profesores()]
        self.app.guardar_ajustes()

    def agregar_profesor(self):
        programa, profesor, cantidad = self.prog.get(), self.prof.get().strip(), self.cant.get().strip()
        if not programa or not profesor:
            messagebox.showwarning("Falta información", "Elija el programa y escriba el nombre del profesor.", parent=self)
            return
        if not cantidad.isdigit() or int(cantidad) < 1:
            messagebox.showwarning("Filas inválidas", "La cantidad de filas debe ser un número mayor que 0.", parent=self)
            return
        actuales = self.profesores()
        if any(p == programa and n.lower() == profesor.lower() for p, n, _ in actuales):
            messagebox.showwarning("Repetido", f"{profesor} ya tiene filas en {programa}.", parent=self)
            return
        usadas = sum(c for p, _, c in actuales if p == programa) + int(cantidad)
        if usadas > (self.filas_validas() or 0):
            messagebox.showwarning("Sin espacio", f"{programa} solo tiene {self.filas.get()} filas; con esta reserva "
                                   f"serían {usadas}. Aumente las filas por programa.", parent=self)
            return
        self.tabla_prof.insert("", "end", values=(programa, profesor, cantidad))
        self.prof.set("")
        self.guardar_preferencias()

    def quitar_profesor(self):
        for i in self.tabla_prof.selection():
            self.tabla_prof.delete(i)
        self.guardar_preferencias()

    def importar_profesores(self):
        ruta = self.app.pedir_abrir("Importar profesores", TIPOS_CSV)
        if not ruta:
            return
        try:
            profesores = onedrive.leer_profesores_csv(ruta)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo importar", str(error), parent=self)
            return
        self.tabla_prof.delete(*self.tabla_prof.get_children())
        for fila in profesores:
            self.tabla_prof.insert("", "end", values=fila)
        self.guardar_preferencias()
        self.registro.escribir(f"Se importaron {len(profesores)} profesores desde {ruta}.")

    def exportar_profesores(self):
        ruta = self.app.pedir_guardar("Exportar profesores", "profesores.csv", ".csv", TIPOS_CSV)
        if not ruta:
            return
        try:
            onedrive.escribir_profesores_csv(ruta, self.profesores())
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo exportar", str(error), parent=self)
            return
        self.registro.escribir(f"Profesores exportados a {ruta}.")

    # --- contraseñas
    def actualizar_contrasenas(self):
        self.tabla_pw.delete(*self.tabla_pw.get_children())
        for (rol, ambito), clave in self.contrasenas.items():
            self.tabla_pw.insert("", "end", iid=f"{rol}\x1f{ambito}",
                                 values=(onedrive.ROLES.get(rol, rol), ambito, clave if self.mostrar else "•" * 10))
        if not self.contrasenas:
            texto, color = "Sin contraseñas cargadas: al generar se crearán nuevas.", "#5F6368"
        elif self.sin_guardar:
            texto, color = f"⚠️ {len(self.contrasenas)} contraseñas SIN GUARDAR. Use «Guardar archivo…».", COLOR_ERROR
        else:
            texto, color = f"{len(self.contrasenas)} contraseñas cargadas y guardadas.", COLOR_ONEDRIVE
        self.estado_pw.configure(text=texto, foreground=color)

    def cargar_contrasenas(self):
        if self.sin_guardar and not messagebox.askyesno(
                "Contraseñas sin guardar", "Las contraseñas actuales no se han guardado y se reemplazarán. ¿Continuar?",
                parent=self):
            return
        ruta = self.app.pedir_abrir("Cargar contraseñas", TIPOS_CSV)
        if not ruta:
            return
        try:
            self.contrasenas = onedrive.leer_contrasenas_csv(ruta)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo cargar", str(error), parent=self)
            return
        self.sin_guardar = False
        self.actualizar_contrasenas()
        self.registro.escribir(f"Se cargaron {len(self.contrasenas)} contraseñas desde {ruta}.")

    def guardar_contrasenas(self):
        """Devuelve True si quedaron guardadas."""
        if not self.contrasenas:
            messagebox.showinfo("Sin contraseñas", "Todavía no hay contraseñas: se crean al generar el Excel.", parent=self)
            return False
        ruta = self.app.pedir_guardar("Guardar contraseñas", onedrive.NOMBRE_SUGERIDO_CONTRASENAS, ".csv", TIPOS_CSV)
        if not ruta:
            return False
        try:
            onedrive.escribir_contrasenas_csv(ruta, self.contrasenas)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo guardar", str(error), parent=self)
            return False
        self.sin_guardar = False
        self.actualizar_contrasenas()
        self.registro.escribir(f"Contraseñas guardadas en {ruta}. Guarde este archivo en un lugar privado.")
        return True

    def alternar_mostrar(self):
        self.mostrar = not self.mostrar
        self.btn_mostrar.configure(text="Ocultar contraseñas" if self.mostrar else "Mostrar contraseñas")
        self.actualizar_contrasenas()

    def clave_seleccionada(self):
        seleccion = self.tabla_pw.selection()
        if not seleccion:
            messagebox.showinfo("Seleccione una fila", "Seleccione una contraseña de la tabla.", parent=self)
            return None
        return tuple(seleccion[0].split("\x1f", 1))

    def copiar_contrasena(self):
        clave = self.clave_seleccionada()
        if clave:
            self.app.copiar(self.contrasenas[clave])
            self.registro.escribir(f"Contraseña de «{clave[1]}» copiada al portapapeles.")

    def regenerar(self, todas):
        if not self.contrasenas:
            messagebox.showinfo("Sin contraseñas", "Cargue un archivo de contraseñas o genere el Excel primero.", parent=self)
            return
        claves = list(self.contrasenas) if todas else [self.clave_seleccionada()]
        if None in claves:
            return
        que = "TODAS las contraseñas" if todas else f"la contraseña de «{claves[0][1]}»"
        if not messagebox.askyesno(
                "Regenerar contraseñas",
                f"Se cambiará {que}. Luego debe generar el Excel de nuevo, reemplazarlo en OneDrive y "
                "entregar las contraseñas nuevas. ¿Continuar?", parent=self):
            return
        for clave in claves:
            self.contrasenas[clave] = onedrive.generar_contrasena()
        self.sin_guardar = True
        self.actualizar_contrasenas()
        self.registro.escribir(f"Se regeneró {que}. Ahora elija dónde guardar el Excel nuevo.", limpiar=True)
        self.generar(limpiar=False)

    # --- generar
    def generar(self, limpiar=True):
        filas = self.filas_validas()
        if not filas:
            messagebox.showwarning("Filas inválidas", f"Las filas deben estar entre {FILAS_MIN} y {FILAS_MAX}.", parent=self)
            return
        if not self.app.plantillas_ok():
            return
        if not self.contrasenas and not messagebox.askyesno(
                "Contraseñas nuevas",
                "No hay contraseñas cargadas, así que se crearán nuevas.\n\nSi ya entregó contraseñas de un Excel "
                "anterior, elija «No» y use «Cargar archivo…» primero. ¿Crear nuevas?", parent=self):
            return
        destino = self.app.pedir_guardar("Guardar Excel para OneDrive", onedrive.NOMBRE_SUGERIDO, ".xlsx", TIPOS_EXCEL)
        if not destino:
            return
        plantillas, profesores, anteriores = self.app.rutas_plantillas(), self.profesores(), dict(self.contrasenas)

        def despues(resultado):
            claves, mensajes = resultado
            for m in mensajes:
                self.registro.escribir(m)
            if claves != anteriores:
                self.contrasenas = claves
                self.sin_guardar = True
            self.actualizar_contrasenas()
            if self.sin_guardar and messagebox.askyesno(
                    "Guardar contraseñas",
                    "Hay contraseñas nuevas o cambiadas. Guárdelas ahora: sin ellas no podrá volver a generar este "
                    "Excel con las mismas claves. ¿Elegir dónde guardarlas?", parent=self):
                self.guardar_contrasenas()

        self.app.ejecutar(lambda: onedrive.generar_excel(plantillas, destino, filas, profesores, anteriores),
                          self.registro, "Excel para OneDrive", despues, limpiar)


# ---------------------------------------------------------------------------
# Pestaña Google Sheets
# ---------------------------------------------------------------------------

class PestanaGoogle(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=(0, 0, 0, 8))
        self.app = app
        conf_gs = app.ajustes["google"]
        try:
            valores = google.validar_config_apps_script(conf_gs["apps_script"])
        except ErrorGeneracion:
            valores = google.config_apps_script_por_defecto()

        encabezado(self, COLOR_GOOGLE, "Plantilla para Google Sheets",
                   "Permisos por cuenta institucional. Genera la plantilla a importar y el Apps Script.").pack(fill="x")
        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)

        # 1. Plantilla
        conf = ttk.LabelFrame(cuerpo, text=" 1. Configuración de la plantilla ", padding=8)
        conf.pack(fill="x")
        fila = ttk.Frame(conf)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(conf_gs["filas"]))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_preferencias())
        nota(conf, "Los usuarios y sus roles no se configuran aquí: se administran en la hoja «Usuarios» "
                   "dentro de Google Sheets.").pack(anchor="w", pady=(6, 0))

        # 2. Apps Script
        script = ttk.LabelFrame(cuerpo, text=" 2. Configuración del Apps Script ", padding=8)
        script.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(script)
        fila.pack(fill="x")
        ttk.Label(fila, text="Dominio institucional:  @").pack(side="left")
        self.dominio = tk.StringVar(value=valores["DOMINIO"])
        entrada = ttk.Entry(fila, textvariable=self.dominio, width=24)
        entrada.pack(side="left")
        entrada.bind("<FocusOut>", lambda _e: self.guardar_preferencias())
        entrada.bind("<Return>", lambda _e: self.guardar_preferencias())
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

        # 3. Generar
        gen = ttk.LabelFrame(cuerpo, text=" 3. Generar y descargar ", padding=8)
        gen.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(gen)
        fila.pack(fill="x")
        self.boton = BotonColor(fila, "Generar plantilla para Google Sheets…", COLOR_GOOGLE, self.generar)
        self.boton.pack(side="left")
        ttk.Button(fila, text="Ver guía de despliegue",
                   command=lambda: app.ver_guia("google-sheets.md", "Guía Google Sheets")).pack(side="left", padx=10)
        fila = ttk.Frame(gen)
        fila.pack(fill="x", pady=(8, 0))
        ttk.Label(fila, text="Apps Script:").pack(side="left")
        ttk.Button(fila, text="Guardar Codigo.gs…", command=lambda: self.guardar_script("Codigo.gs")).pack(side="left", padx=(6, 0))
        ttk.Button(fila, text="Guardar appsscript.json…", command=lambda: self.guardar_script("appsscript.json")).pack(side="left", padx=6)
        ttk.Button(fila, text="Copiar Codigo.gs", command=lambda: self.copiar_script("Codigo.gs")).pack(side="left")
        ttk.Button(fila, text="Copiar appsscript.json", command=lambda: self.copiar_script("appsscript.json")).pack(side="left", padx=6)
        nota(gen, "Importe la plantilla a Google Sheets y pegue el código en Extensiones → Apps Script "
                  "(pasos completos en la guía). Si cambia la configuración, vuelva a copiar Codigo.gs.").pack(anchor="w", pady=(6, 0))

        self.registro = Registro(cuerpo, "Resultado (Google Sheets):")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))

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
        self.estado_script.configure(text="✓ Configuración guardada.", foreground=COLOR_GOOGLE)
        return valores

    def filas_validas(self):
        try:
            n = int(self.filas.get())
        except ValueError:
            return None
        return n if FILAS_MIN <= n <= FILAS_MAX else None

    def guardar_preferencias(self):
        if self.filas_validas():
            self.app.ajustes["google"]["filas"] = self.filas_validas()
        valores = self.config_valida()
        if valores:
            self.app.ajustes["google"]["apps_script"] = valores
        self.app.guardar_ajustes()

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
        ruta = self.app.pedir_guardar(f"Guardar {nombre}", nombre, Path(nombre).suffix, tipos)
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
            self.app.copiar(texto)
            self.registro.escribir(f"{nombre} copiado al portapapeles. Péguelo en Extensiones → Apps Script.")

    def generar(self):
        filas, valores = self.filas_validas(), self.config_valida()
        if not filas:
            messagebox.showwarning("Filas inválidas", f"Las filas deben estar entre {FILAS_MIN} y {FILAS_MAX}.", parent=self)
            return
        if not valores or not self.app.plantillas_ok():
            return
        destino = self.app.pedir_guardar("Guardar plantilla para Google Sheets", google.NOMBRE_SUGERIDO, ".xlsx", TIPOS_EXCEL)
        if not destino:
            return
        plantillas = self.app.rutas_plantillas()

        def despues(mensajes):
            for m in mensajes:
                self.registro.escribir(m)
            self.registro.escribir("Siguiente: guarde o copie Codigo.gs y appsscript.json (botones de arriba).")

        self.app.ejecutar(lambda: google.generar_plantilla(plantillas, destino, filas, valores),
                          self.registro, "plantilla para Google Sheets", despues)


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Cronograma 2027 Bloque Neón — Generador {__version__}")
        self.ocupado = False
        self.ajustes = preferencias.leer()
        self.programas = []

        # Plantillas de entrada (comunes a las dos versiones)
        arriba = ttk.LabelFrame(self, text=" Plantillas de entrada (comunes a las dos versiones) ", padding=8)
        arriba.pack(fill="x", padx=12, pady=(10, 6))
        self.estado_plantillas = ttk.Label(arriba, text="")
        self.estado_plantillas.pack(side="left")
        ttk.Button(arriba, text="Ver lista", command=self.ver_plantillas).pack(side="right")
        ttk.Button(arriba, text="Usar las incluidas", command=self.usar_incluidas).pack(side="right", padx=6)
        ttk.Button(arriba, text="Elegir carpeta…", command=self.elegir_carpeta).pack(side="right")
        ttk.Button(arriba, text="Elegir archivos…", command=self.elegir_archivos).pack(side="right", padx=6)

        self.pestanas = ttk.Notebook(self)
        self.pestanas.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.onedrive = PestanaOneDrive(self, self.pestanas)
        self.google = PestanaGoogle(self, self.pestanas)
        self.pestanas.add(self.onedrive, text="   Excel para OneDrive   ")
        self.pestanas.add(self.google, text="   Google Sheets   ")
        self.cargar_plantillas()

        self.protocol("WM_DELETE_WINDOW", self.cerrar)
        self.update_idletasks()
        alto = min(self.winfo_reqheight(), self.winfo_screenheight() - 80)
        self.geometry(f"{max(self.winfo_reqwidth(), 900)}x{alto}")
        self.minsize(860, min(alto, 640))

    # --- plantillas
    def rutas_plantillas(self):
        elegidas = self.ajustes["plantillas"]
        if elegidas:
            return [Path(p) for p in elegidas]
        incluidas = carpeta_plantillas_incluidas()
        return [incluidas] if incluidas else []

    def cargar_plantillas(self):
        """Lee las plantillas elegidas para validar y obtener los programas."""
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
        self.onedrive.combo_prog.configure(values=self.programas)

    def plantillas_ok(self):
        if not self.programas:
            messagebox.showwarning("Plantillas", "Primero elija plantillas válidas (arriba).", parent=self)
            return False
        return True

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
            from ..comun.plantillas import expandir_rutas, programa_de
            lineas = [f"{programa_de(a):<16} {a}" for a in expandir_rutas(self.rutas_plantillas())]
        except ErrorGeneracion as error:
            lineas = [str(error)]
        mostrar_texto(self, "Plantillas de entrada", "\n".join(lineas))

    # --- diálogos de archivos (la última carpeta usada se recuerda, nada más)
    def carpeta_inicial(self):
        carpeta = self.ajustes.get("ultima_carpeta")
        return carpeta if carpeta and Path(carpeta).is_dir() else str(Path.home())

    def recordar_carpeta(self, ruta):
        self.ajustes["ultima_carpeta"] = str(Path(ruta).parent)
        self.guardar_ajustes()

    def pedir_abrir(self, titulo, tipos):
        ruta = filedialog.askopenfilename(parent=self, title=titulo, filetypes=tipos, initialdir=self.carpeta_inicial())
        if ruta:
            self.recordar_carpeta(ruta)
        return ruta or None

    def pedir_guardar(self, titulo, nombre, extension, tipos):
        ruta = filedialog.asksaveasfilename(parent=self, title=titulo, initialfile=nombre, defaultextension=extension,
                                            filetypes=tipos, initialdir=self.carpeta_inicial())
        if ruta:
            self.recordar_carpeta(ruta)
        return ruta or None

    # --- utilidades
    def guardar_ajustes(self):
        preferencias.guardar(self.ajustes)

    def copiar(self, texto):
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.update()

    def ver_guia(self, archivo, titulo):
        try:
            texto = (GUIAS / archivo).read_text(encoding="utf-8")
        except OSError:
            texto = "No se encontró la guía en esta instalación."
        mostrar_texto(self, titulo, texto)

    def ejecutar(self, funcion, registro, nombre, al_terminar, limpiar=True):
        """Corre `funcion` en segundo plano; el resultado vuelve al hilo de la ventana mediante una cola."""
        if self.ocupado:
            return
        self.ocupado = True
        for pestana in (self.onedrive, self.google):
            pestana.boton.habilitar(False)
        registro.escribir(f"Generando {nombre}…", limpiar=limpiar)
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
            for pestana in (self.onedrive, self.google):
                pestana.boton.habilitar(True)
            if estado == "ok":
                al_terminar(valor)
                registro.escribir("✓ Listo.")
            else:
                registro.escribir(f"✗ {valor}")
                messagebox.showerror("No se pudo generar", valor if estado == "error" else
                                     "Ocurrió un error inesperado. El detalle quedó en el cuadro de resultado.", parent=self)

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(100, revisar)

    def cerrar(self):
        if self.onedrive.sin_guardar:
            respuesta = messagebox.askyesnocancel(
                "Contraseñas sin guardar",
                "Hay contraseñas que no se han guardado en un archivo. Si cierra, se perderán y el Excel generado "
                "quedará con claves que nadie conoce.\n\n¿Guardarlas antes de salir?", parent=self)
            if respuesta is None or (respuesta and not self.onedrive.guardar_contrasenas()):
                return
        self.destroy()


def ejecutar():
    App().mainloop()
