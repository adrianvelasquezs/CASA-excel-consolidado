"""
Interfaz gráfica. El libro se genera desde cero con la configuración que se
edita aquí: programas, columnas y quién puede editar cada columna. No usa
rutas fijas: la configuración se importa/exporta y cada resultado se guarda
donde el usuario elija.
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
from ..comun.configuracion import (
    ANCHO_MAX,
    ANCHO_MIN,
    FILAS_MAX,
    FILAS_MIN,
    TIPOS,
    Columna,
    Configuracion,
    ErrorGeneracion,
    predeterminada,
    validar_columna,
    validar_programa,
)
from ..google_sheets import generador as google

COLOR = "#188038"
COLOR_ERROR = "#C5221F"
GUIA = Path(__file__).resolve().parent.parent / "guias" / "google-sheets.md"
TIPOS_EXCEL = [("Libro de Excel", "*.xlsx")]
TIPOS_JSON = [("Configuración (JSON)", "*.json")]
TIPO_POR_ETIQUETA = {etiqueta: clave for clave, etiqueta in TIPOS.items()}


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


def nota(padre, texto="", ancho=880):
    return ttk.Label(padre, text=texto, foreground="#5F6368", wraplength=ancho, justify="left")


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
    """Ventana de solo lectura (guía)."""
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
# Pestaña «Columnas y permisos»
# ---------------------------------------------------------------------------

class PestanaColumnas(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=10)
        self.app = app

        nota(self, "Columnas de cada hoja de programa, en orden. «Programa», «Registrado por» y «Fecha de registro» "
                   "las agrega el sistema. ADMIN siempre puede editar todo.").pack(anchor="w", pady=(0, 6))
        arriba = ttk.Frame(self)
        arriba.pack(fill="both", expand=True)
        self.tabla = ttk.Treeview(arriba, columns=("nombre", "tipo", "editan", "identifica"), show="headings",
                                  height=9, selectmode="browse")
        for ident, titulo, ancho in (("nombre", "Columna", 380), ("tipo", "Tipo", 170),
                                     ("editan", "Quién puede editar", 200), ("identifica", "Identifica", 80)):
            self.tabla.heading(ident, text=titulo)
            self.tabla.column(ident, width=ancho, anchor="w" if ident != "identifica" else "center")
        self.tabla.pack(side="left", fill="both", expand=True)
        barra = ttk.Scrollbar(arriba, orient="vertical", command=self.tabla.yview)
        barra.pack(side="left", fill="y")
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self.cargar_formulario())
        lado = ttk.Frame(arriba, padding=(8, 0, 0, 0))
        lado.pack(side="left", fill="y")
        for texto, comando in (("↑ Subir", lambda: self.mover(-1)), ("↓ Bajar", lambda: self.mover(1)),
                               ("Quitar", self.quitar)):
            ttk.Button(lado, text=texto, command=comando, width=12).pack(pady=2)

        # Formulario
        form = ttk.LabelFrame(self, text=" Columna seleccionada / nueva ", padding=8)
        form.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(form)
        fila.pack(fill="x")
        ttk.Label(fila, text="Nombre:").pack(side="left")
        self.nombre = tk.StringVar()
        ttk.Entry(fila, textvariable=self.nombre, width=46).pack(side="left", padx=(4, 12))
        ttk.Label(fila, text="Tipo:").pack(side="left")
        self.tipo = tk.StringVar(value=TIPOS["texto"])
        combo = ttk.Combobox(fila, textvariable=self.tipo, values=list(TIPOS.values()), state="readonly", width=20)
        combo.pack(side="left", padx=(4, 12))
        combo.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_opciones())
        ttk.Label(fila, text="Ancho:").pack(side="left")
        self.ancho = tk.StringVar(value="20")
        ttk.Spinbox(fila, from_=ANCHO_MIN, to=ANCHO_MAX, width=5, textvariable=self.ancho).pack(side="left", padx=4)

        fila = ttk.Frame(form)
        fila.pack(fill="x", pady=(6, 0))
        ttk.Label(fila, text="Opciones de la lista (separadas por coma):").pack(side="left")
        self.opciones = tk.StringVar()
        self.entrada_opciones = ttk.Entry(fila, textvariable=self.opciones, width=50)
        self.entrada_opciones.pack(side="left", padx=4)

        fila = ttk.Frame(form)
        fila.pack(fill="x", pady=(6, 0))
        ttk.Label(fila, text="Quién puede editar:").pack(side="left")
        self.profesores, self.gestores, self.casa = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
        self.identifica = tk.BooleanVar()
        for texto, var in (("Profesores", self.profesores), ("Gestores", self.gestores), ("CASA", self.casa)):
            ttk.Checkbutton(fila, text=texto, variable=var).pack(side="left", padx=(8, 0))
        ttk.Label(fila, text="   ").pack(side="left")
        ttk.Checkbutton(fila, text="Identifica el curso", variable=self.identifica).pack(side="left")

        fila = ttk.Frame(form)
        fila.pack(fill="x", pady=(8, 0))
        ttk.Button(fila, text="Guardar cambios de la columna", command=self.guardar).pack(side="left")
        ttk.Button(fila, text="Agregar como columna nueva", command=self.agregar).pack(side="left", padx=6)
        ttk.Button(fila, text="Limpiar", command=self.limpiar).pack(side="left")
        nota(form, "«Identifica el curso»: la fila cuenta como registrada cuando alguna de estas columnas tiene datos "
                   "(se usa en la columna Programa y en el Consolidado).").pack(anchor="w", pady=(6, 0))
        self.refrescar()
        self.actualizar_opciones()

    @property
    def columnas(self):
        return self.app.configuracion.columnas

    def refrescar(self, seleccionar=None):
        self.tabla.delete(*self.tabla.get_children())
        for i, col in enumerate(self.columnas):
            tipo = TIPOS[col.tipo] + (f": {', '.join(col.opciones)}" if col.tipo == "lista" else "")
            self.tabla.insert("", "end", iid=str(i), values=(col.nombre, tipo, col.quien_edita(),
                                                             "✓" if col.identifica else ""))
        if seleccionar is not None and 0 <= seleccionar < len(self.columnas):
            self.tabla.selection_set(str(seleccionar))
            self.tabla.see(str(seleccionar))

    def seleccion(self):
        sel = self.tabla.selection()
        return int(sel[0]) if sel else None

    def actualizar_opciones(self):
        es_lista = TIPO_POR_ETIQUETA.get(self.tipo.get()) == "lista"
        self.entrada_opciones.configure(state="normal" if es_lista else "disabled")

    def cargar_formulario(self):
        i = self.seleccion()
        if i is None:
            return
        col = self.columnas[i]
        self.nombre.set(col.nombre)
        self.tipo.set(TIPOS[col.tipo])
        self.ancho.set(str(col.ancho))
        self.opciones.set(", ".join(col.opciones))
        for var, valor in ((self.profesores, col.profesores), (self.gestores, col.gestores),
                           (self.casa, col.casa), (self.identifica, col.identifica)):
            var.set(valor)
        self.actualizar_opciones()

    def limpiar(self):
        self.tabla.selection_remove(*self.tabla.selection())
        self.nombre.set("")
        self.tipo.set(TIPOS["texto"])
        self.ancho.set("20")
        self.opciones.set("")
        for var in (self.profesores, self.gestores, self.casa, self.identifica):
            var.set(False)
        self.actualizar_opciones()

    def columna_del_formulario(self, excepto=None):
        """Columna con los valores del formulario, o None (con aviso) si algo está mal."""
        tipo = TIPO_POR_ETIQUETA.get(self.tipo.get(), "texto")
        try:
            ancho = int(self.ancho.get())
        except ValueError:
            ancho = -1
        col = Columna(
            nombre=self.nombre.get().strip(), tipo=tipo, ancho=ancho,
            opciones=[o.strip() for o in self.opciones.get().split(",") if o.strip()] if tipo == "lista" else [],
            profesores=self.profesores.get(), gestores=self.gestores.get(), casa=self.casa.get(),
            identifica=self.identifica.get(),
        )
        try:
            validar_columna(col)
            if any(c.nombre.lower() == col.nombre.lower() for j, c in enumerate(self.columnas) if j != excepto):
                raise ErrorGeneracion(f"Ya existe una columna llamada «{col.nombre}».")
        except ErrorGeneracion as error:
            messagebox.showwarning("Columna inválida", str(error), parent=self)
            return None
        return col

    def guardar(self):
        i = self.seleccion()
        if i is None:
            messagebox.showinfo("Seleccione una columna", "Seleccione en la tabla la columna que quiere cambiar, "
                                "o use «Agregar como columna nueva».", parent=self)
            return
        col = self.columna_del_formulario(excepto=i)
        if col:
            self.columnas[i] = col
            self.app.configuracion_cambiada(f"Columna «{col.nombre}» actualizada.")
            self.refrescar(seleccionar=i)

    def agregar(self):
        col = self.columna_del_formulario()
        if not col:
            return
        i = self.seleccion()
        posicion = i + 1 if i is not None else len(self.columnas)
        self.columnas.insert(posicion, col)
        self.app.configuracion_cambiada(f"Columna «{col.nombre}» agregada.")
        self.refrescar(seleccionar=posicion)

    def quitar(self):
        i = self.seleccion()
        if i is None:
            return
        nombre = self.columnas[i].nombre
        if not messagebox.askyesno("Quitar columna", f"¿Quitar la columna «{nombre}»?", parent=self):
            return
        del self.columnas[i]
        self.app.configuracion_cambiada(f"Columna «{nombre}» quitada.")
        self.refrescar(seleccionar=min(i, len(self.columnas) - 1))

    def mover(self, paso):
        i = self.seleccion()
        if i is None or not 0 <= i + paso < len(self.columnas):
            return
        self.columnas[i], self.columnas[i + paso] = self.columnas[i + paso], self.columnas[i]
        self.app.configuracion_cambiada()
        self.refrescar(seleccionar=i + paso)


# ---------------------------------------------------------------------------
# Pestaña «Programas»
# ---------------------------------------------------------------------------

class PestanaProgramas(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=10)
        self.app = app
        nota(self, "Cada programa es una hoja del libro. El nombre es el que se usa en la columna «Programa(s)» "
                   "de la hoja Usuarios (máximo 31 caracteres).").pack(anchor="w", pady=(0, 6))
        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True)
        self.lista = tk.Listbox(cuerpo, height=12, exportselection=False, activestyle="none")
        self.lista.pack(side="left", fill="both", expand=True)
        barra = ttk.Scrollbar(cuerpo, orient="vertical", command=self.lista.yview)
        barra.pack(side="left", fill="y")
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.bind("<<ListboxSelect>>", lambda _e: self.cargar())
        lado = ttk.Frame(cuerpo, padding=(8, 0, 0, 0))
        lado.pack(side="left", fill="y")
        ttk.Label(lado, text="Nombre:").pack(anchor="w")
        self.nombre = tk.StringVar()
        entrada = ttk.Entry(lado, textvariable=self.nombre, width=24)
        entrada.pack(anchor="w", pady=(0, 6))
        entrada.bind("<Return>", lambda _e: self.agregar())
        for texto, comando in (("Agregar", self.agregar), ("Renombrar seleccionado", self.renombrar),
                               ("Quitar seleccionado", self.quitar), ("↑ Subir", lambda: self.mover(-1)),
                               ("↓ Bajar", lambda: self.mover(1))):
            ttk.Button(lado, text=texto, command=comando, width=22).pack(anchor="w", pady=2)
        self.refrescar()

    @property
    def programas(self):
        return self.app.configuracion.programas

    def refrescar(self, seleccionar=None):
        self.lista.delete(0, "end")
        for programa in self.programas:
            self.lista.insert("end", programa)
        if seleccionar is not None and 0 <= seleccionar < len(self.programas):
            self.lista.selection_set(seleccionar)
            self.lista.see(seleccionar)

    def seleccion(self):
        sel = self.lista.curselection()
        return sel[0] if sel else None

    def cargar(self):
        i = self.seleccion()
        if i is not None:
            self.nombre.set(self.programas[i])

    def nombre_valido(self, excepto=None):
        nombre = self.nombre.get().strip()
        try:
            validar_programa(nombre)
            if any(p.lower() == nombre.lower() for j, p in enumerate(self.programas) if j != excepto):
                raise ErrorGeneracion(f"El programa «{nombre}» ya existe.")
        except ErrorGeneracion as error:
            messagebox.showwarning("Programa inválido", str(error), parent=self)
            return None
        return nombre

    def agregar(self):
        nombre = self.nombre_valido()
        if nombre:
            self.programas.append(nombre)
            self.app.configuracion_cambiada(f"Programa «{nombre}» agregado.")
            self.refrescar(seleccionar=len(self.programas) - 1)
            self.nombre.set("")

    def renombrar(self):
        i = self.seleccion()
        if i is None:
            return
        nombre = self.nombre_valido(excepto=i)
        if nombre:
            anterior, self.programas[i] = self.programas[i], nombre
            self.app.configuracion_cambiada(f"Programa «{anterior}» renombrado a «{nombre}».")
            self.refrescar(seleccionar=i)

    def quitar(self):
        i = self.seleccion()
        if i is None:
            return
        nombre = self.programas[i]
        if messagebox.askyesno("Quitar programa", f"¿Quitar el programa «{nombre}»?", parent=self):
            del self.programas[i]
            self.app.configuracion_cambiada(f"Programa «{nombre}» quitado.")
            self.refrescar(seleccionar=min(i, len(self.programas) - 1))

    def mover(self, paso):
        i = self.seleccion()
        if i is None or not 0 <= i + paso < len(self.programas):
            return
        self.programas[i], self.programas[i + paso] = self.programas[i + paso], self.programas[i]
        self.app.configuracion_cambiada()
        self.refrescar(seleccionar=i + paso)


# ---------------------------------------------------------------------------
# Pestaña «Opciones»
# ---------------------------------------------------------------------------

class PestanaOpciones(ttk.Frame):
    def __init__(self, app, padre, valores):
        super().__init__(padre, padding=10)
        self.app = app
        marco = ttk.LabelFrame(self, text=" Plantilla ", padding=8)
        marco.pack(fill="x")
        fila = ttk.Frame(marco)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(app.configuracion.filas))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_filas())
        nota(marco, "Los usuarios y sus roles no se configuran aquí: se administran en la hoja «Usuarios» "
                    "dentro de Google Sheets.").pack(anchor="w", pady=(6, 0))

        script = ttk.LabelFrame(self, text=" Apps Script ", padding=8)
        script.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(script)
        fila.pack(fill="x")
        ttk.Label(fila, text="Dominio institucional:  @").pack(side="left")
        self.dominio = tk.StringVar(value=valores["DOMINIO"])
        caja = ttk.Entry(fila, textvariable=self.dominio, width=24)
        caja.pack(side="left")
        caja.bind("<FocusOut>", lambda _e: self.guardar_script())
        caja.bind("<Return>", lambda _e: self.guardar_script())
        ttk.Label(fila, text="      Restaurar permisos cada día a las").pack(side="left")
        self.hora = tk.StringVar(value=str(valores["HORA_RESTAURACION"]))
        ttk.Spinbox(fila, from_=0, to=23, width=4, textvariable=self.hora, wrap=True,
                    command=self.guardar_script).pack(side="left", padx=4)
        ttk.Label(fila, text="h").pack(side="left")
        self.gestores = tk.BooleanVar(value=valores["GESTOR_EDITA_FILAS_DE_PROFESORES"])
        self.alertas = tk.BooleanVar(value=valores["ALERTAS_POR_CORREO"])
        ttk.Checkbutton(script, text="Los gestores pueden editar las filas registradas por sus profesores",
                        variable=self.gestores, command=self.guardar_script).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(script, text="Avisar por correo a los ADMIN cuando alguien edita algo que no le corresponde",
                        variable=self.alertas, command=self.guardar_script).pack(anchor="w")
        self.estado = nota(script)
        self.estado.pack(anchor="w", pady=(4, 0))
        nota(script, "Si cambia esto con el libro ya en uso, vuelva a copiar Codigo.gs y péguelo en Extensiones → "
                     "Apps Script.").pack(anchor="w")

    def guardar_filas(self):
        try:
            filas = int(self.filas.get())
        except ValueError:
            return
        if FILAS_MIN <= filas <= FILAS_MAX and filas != self.app.configuracion.filas:
            self.app.configuracion.filas = filas
            self.app.configuracion_cambiada()

    def config_script(self):
        return {"DOMINIO": self.dominio.get(), "HORA_RESTAURACION": self.hora.get(),
                "GESTOR_EDITA_FILAS_DE_PROFESORES": self.gestores.get(), "ALERTAS_POR_CORREO": self.alertas.get()}

    def config_valida(self):
        try:
            valores = google.validar_config_apps_script(self.config_script())
        except ErrorGeneracion as error:
            self.estado.configure(text=f"⚠️ {error}", foreground=COLOR_ERROR)
            return None
        self.estado.configure(text="✓ Opciones guardadas.", foreground=COLOR)
        return valores

    def guardar_script(self):
        valores = self.config_valida()
        if valores:
            self.app.ajustes["google"]["apps_script"] = valores
            self.app.guardar_ajustes()
        return valores


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Cronograma 2027 Bloque Neón — Generador {__version__}")
        self.ocupado = False
        self.ajustes = preferencias.leer()
        self.configuracion, aviso = self.cargar_configuracion_guardada()
        try:
            valores = google.validar_config_apps_script(self.ajustes["google"]["apps_script"])
        except ErrorGeneracion:
            valores = google.config_apps_script_por_defecto()

        marco = tk.Frame(self, bg=COLOR, padx=14, pady=8)
        marco.pack(fill="x")
        tk.Label(marco, text="Cronograma consolidado para Google Sheets", bg=COLOR, fg="white",
                 font=tkfont.Font(size=17, weight="bold")).pack(anchor="w")
        tk.Label(marco, text="Defina programas, columnas y permisos; genere la plantilla y el Apps Script.",
                 bg=COLOR, fg="white").pack(anchor="w")
        cuerpo = ttk.Frame(self, padding=(12, 8, 12, 12))
        cuerpo.pack(fill="both", expand=True)

        barra = ttk.Frame(cuerpo)
        barra.pack(fill="x", pady=(0, 6))
        self.estado_config = ttk.Label(barra, text="")
        self.estado_config.pack(side="left")
        ttk.Button(barra, text="Restablecer predeterminada", command=self.restablecer).pack(side="right")
        ttk.Button(barra, text="Exportar configuración…", command=self.exportar).pack(side="right", padx=6)
        ttk.Button(barra, text="Importar configuración…", command=self.importar).pack(side="right")

        self.pestanas = ttk.Notebook(cuerpo)
        self.pestanas.pack(fill="both", expand=True)
        self.columnas = PestanaColumnas(self, self.pestanas)
        self.programas = PestanaProgramas(self, self.pestanas)
        self.opciones = PestanaOpciones(self, self.pestanas, valores)
        self.pestanas.add(self.columnas, text="  Columnas y permisos  ")
        self.pestanas.add(self.programas, text="  Programas  ")
        self.pestanas.add(self.opciones, text="  Opciones  ")

        gen = ttk.Frame(cuerpo)
        gen.pack(fill="x", pady=(10, 0))
        self.boton = BotonColor(gen, "Generar plantilla para Google Sheets…", COLOR, self.generar)
        self.boton.pack(side="left")
        lado = ttk.Frame(gen)
        lado.pack(side="left", padx=(12, 0))
        fila = ttk.Frame(lado)
        fila.pack(anchor="w")
        ttk.Button(fila, text="Guardar Codigo.gs…", command=lambda: self.guardar_script("Codigo.gs")).pack(side="left")
        ttk.Button(fila, text="Copiar Codigo.gs", command=lambda: self.copiar_script("Codigo.gs")).pack(side="left", padx=6)
        ttk.Button(fila, text="Ver guía de despliegue", command=self.ver_guia).pack(side="left")
        fila = ttk.Frame(lado)
        fila.pack(anchor="w", pady=(4, 0))
        ttk.Button(fila, text="Guardar appsscript.json…", command=lambda: self.guardar_script("appsscript.json")).pack(side="left")
        ttk.Button(fila, text="Copiar appsscript.json", command=lambda: self.copiar_script("appsscript.json")).pack(side="left", padx=6)

        self.registro = Registro(cuerpo, "Resultado:")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))
        if aviso:
            self.registro.escribir(aviso)
        self.actualizar_estado()

        self.update_idletasks()
        alto = min(self.winfo_reqheight(), self.winfo_screenheight() - 80)
        self.geometry(f"{max(self.winfo_reqwidth(), 940)}x{alto}")
        self.minsize(900, min(alto, 640))

    # --- configuración
    def cargar_configuracion_guardada(self):
        guardada = self.ajustes.get("configuracion")
        if not guardada:
            return predeterminada(), None
        try:
            return Configuracion.desde_dict(guardada).validar(), None
        except ErrorGeneracion as error:
            return predeterminada(), f"⚠️ La configuración guardada no era válida ({error}); se cargó la predeterminada."

    def configuracion_cambiada(self, mensaje=None):
        self.ajustes["configuracion"] = self.configuracion.a_dict()
        self.guardar_ajustes()
        self.actualizar_estado()
        if mensaje:
            self.registro.escribir(mensaje)

    def actualizar_estado(self):
        resumen = f"{len(self.configuracion.programas)} programas · {len(self.configuracion.columnas)} columnas · {self.configuracion.filas} filas"
        try:
            self.configuracion.validar()
            self.estado_config.configure(text=f"Configuración: {resumen}  ✓", foreground="")
        except ErrorGeneracion as error:
            self.estado_config.configure(text=f"⚠️ {error}", foreground=COLOR_ERROR)

    def aplicar_configuracion(self, configuracion, mensaje):
        self.configuracion = configuracion
        self.columnas.refrescar()
        self.columnas.limpiar()
        self.programas.refrescar()
        self.opciones.filas.set(str(configuracion.filas))
        self.configuracion_cambiada(mensaje)

    def importar(self):
        ruta = filedialog.askopenfilename(parent=self, title="Importar configuración", filetypes=TIPOS_JSON,
                                          initialdir=self.carpeta_inicial())
        if not ruta:
            return
        self.recordar_carpeta(ruta)
        try:
            configuracion = Configuracion.leer_json(ruta)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo importar", str(error), parent=self)
            return
        self.aplicar_configuracion(configuracion, f"Configuración importada desde {ruta}.")

    def exportar(self):
        ruta = self.pedir_guardar("Exportar configuración", "configuracion-cronograma.json", ".json", TIPOS_JSON)
        if not ruta:
            return
        try:
            self.configuracion.guardar_json(ruta)
        except ErrorGeneracion as error:
            messagebox.showerror("No se pudo exportar", str(error), parent=self)
            return
        self.registro.escribir(f"Configuración exportada a {ruta}.")

    def restablecer(self):
        if messagebox.askyesno("Restablecer", "Se reemplazarán los programas, columnas y permisos actuales por los "
                               "predeterminados. ¿Continuar?", parent=self):
            self.aplicar_configuracion(predeterminada(), "Se restableció la configuración predeterminada.")

    # --- Apps Script
    def texto_script(self, nombre):
        if nombre == "appsscript.json":
            return google.manifiesto_apps_script()
        valores = self.opciones.config_valida()
        if not valores:
            self.pestanas.select(self.opciones)
            return None
        return google.codigo_apps_script(valores)

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
        try:
            self.configuracion.validar()
        except ErrorGeneracion as error:
            messagebox.showwarning("Configuración incompleta", str(error), parent=self)
            return
        valores = self.opciones.config_valida()
        if not valores:
            self.pestanas.select(self.opciones)
            return
        destino = self.pedir_guardar("Guardar plantilla para Google Sheets", google.NOMBRE_SUGERIDO, ".xlsx", TIPOS_EXCEL)
        if not destino:
            return
        configuracion = self.configuracion.copia()

        def despues(mensajes):
            for m in mensajes:
                self.registro.escribir(m)
            self.registro.escribir("Siguiente: guarde o copie Codigo.gs y appsscript.json.")

        self.ejecutar(lambda: google.generar_plantilla(configuracion, destino, valores), despues)

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
