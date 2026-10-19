"""
Interfaz gráfica del Cronograma 2027 Bloque Neón.

Dos pestañas totalmente separadas:
  - Excel para OneDrive: configura onedrive/config/, administra las contraseñas
    y genera onedrive/salida/.
  - Google Sheets: configura google-sheets/config/ y el Apps Script, genera
    google-sheets/salida/ y copia el código para pegarlo en Apps Script.

Uso (desde la raíz del proyecto):
  uv run interfaz.py
o doble clic en «Abrir interfaz.command» (Mac) / «Abrir interfaz.bat» (Windows).
"""

import csv
import importlib.util
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "comun"))

from plantillas import RAW_DIR, guardar_ajustes, leer_ajustes, leer_plantillas  # noqa: E402

ONEDRIVE = ROOT / "onedrive"
GOOGLE = ROOT / "google-sheets"

COLOR_ONEDRIVE = "#0F6CBD"
COLOR_GOOGLE = "#188038"
FILAS_MIN, FILAS_MAX = 10, 2000

ROLES_CONTRASENA = {"admin": "Administrador", "casa": "Equipo CASA", "programa": "Programa", "profesor": "Profesor"}

# Valores editables de CONFIG en google-sheets/apps-script/Codigo.gs
CAMPOS_SCRIPT = {
    "DOMINIO": "texto",
    "GESTOR_EDITA_FILAS_DE_PROFESORES": "bool",
    "ALERTAS_POR_CORREO": "bool",
    "HORA_RESTAURACION": "hora",
}


def cargar_modulo(nombre, ruta):
    spec = importlib.util.spec_from_file_location(nombre, ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


# El generador de OneDrive es la única fuente de verdad para rutas y contraseñas.
gen_onedrive = cargar_modulo("generar_excel", ONEDRIVE / "generar_excel.py")
gen_google = cargar_modulo("generar_plantilla", GOOGLE / "generar_plantilla.py")
CODIGO_GS = GOOGLE / "apps-script" / "Codigo.gs"
MANIFIESTO = GOOGLE / "apps-script" / "appsscript.json"


def abrir(ruta):
    """Abre un archivo o carpeta con la aplicación predeterminada del sistema."""
    ruta = str(ruta)
    if sys.platform == "darwin":
        subprocess.Popen(["open", ruta])
    elif os.name == "nt":
        os.startfile(ruta)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", ruta])


def leer_config_script():
    texto = CODIGO_GS.read_text(encoding="utf-8")
    valores = {}
    for clave, tipo in CAMPOS_SCRIPT.items():
        m = re.search(rf"^\s*{clave}:\s*(.+?),\s*$", texto, re.M)
        if not m:
            raise ValueError(f"No se encontró {clave} en Codigo.gs")
        crudo = m.group(1).strip()
        if tipo == "texto":
            valores[clave] = crudo.strip("'\"")
        elif tipo == "bool":
            valores[clave] = crudo == "true"
        else:
            valores[clave] = int(crudo)
    return valores


def guardar_config_script(valores):
    texto = CODIGO_GS.read_text(encoding="utf-8")
    for clave, tipo in CAMPOS_SCRIPT.items():
        valor = valores[clave]
        if tipo == "texto":
            literal = "'" + str(valor).replace("\\", "\\\\").replace("'", "\\'") + "'"
        elif tipo == "bool":
            literal = "true" if valor else "false"
        else:
            literal = str(int(valor))
        texto, n = re.subn(rf"^(\s*{clave}:\s*).+?(,\s*)$", lambda m: m.group(1) + literal + m.group(2),
                           texto, count=1, flags=re.M)
        if n != 1:
            raise ValueError(f"No se pudo actualizar {clave} en Codigo.gs")
    CODIGO_GS.write_text(texto, encoding="utf-8")


def leer_profesores():
    if not gen_onedrive.ARCHIVO_PROFESORES.exists():
        return []
    with open(gen_onedrive.ARCHIVO_PROFESORES, newline="", encoding="utf-8-sig") as f:
        return [(r.get("programa", "").strip(), r.get("profesor", "").strip(), r.get("filas", "").strip())
                for r in csv.DictReader(f) if (r.get("programa") or r.get("profesor"))]


def guardar_profesores(filas):
    gen_onedrive.ARCHIVO_PROFESORES.parent.mkdir(parents=True, exist_ok=True)
    with open(gen_onedrive.ARCHIVO_PROFESORES, "w", newline="", encoding="utf-8") as f:
        escritor = csv.writer(f)
        escritor.writerow(["programa", "profesor", "filas"])
        escritor.writerows(filas)


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
    marco = tk.Frame(padre, bg=color, padx=14, pady=10)
    tk.Label(marco, text=titulo, bg=color, fg="white", font=tkfont.Font(size=17, weight="bold")).pack(anchor="w")
    tk.Label(marco, text=subtitulo, bg=color, fg="white", justify="left").pack(anchor="w")
    return marco


def nota(padre, texto):
    return ttk.Label(padre, text=texto, foreground="#5F6368", wraplength=880, justify="left")


class Registro(ttk.Frame):
    """Caja de texto de solo lectura donde se muestra lo que imprime el generador."""

    def __init__(self, padre, titulo):
        super().__init__(padre)
        ttk.Label(self, text=titulo).pack(anchor="w")
        self.texto = tk.Text(self, height=4, wrap="word", state="disabled", relief="solid", borderwidth=1)
        self.texto.pack(fill="both", expand=True)

    def escribir(self, mensaje, limpiar=False):
        self.texto.configure(state="normal")
        if limpiar:
            self.texto.delete("1.0", "end")
        self.texto.insert("end", mensaje.rstrip() + "\n")
        self.texto.see("end")
        self.texto.configure(state="disabled")


# ---------------------------------------------------------------------------
# Pestaña OneDrive
# ---------------------------------------------------------------------------

class PestanaOneDrive(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=(0, 0, 0, 8))
        self.app = app
        self.mostrar = False

        encabezado(self, COLOR_ONEDRIVE, "Excel para OneDrive / SharePoint",
                   "Permisos con contraseñas.   Entradas: raw/ + onedrive/config/   ·   "
                   "Salida: onedrive/salida/").pack(fill="x")
        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)

        # 1. Configuración
        conf = ttk.LabelFrame(cuerpo, text=" 1. Configuración ", padding=8)
        conf.pack(fill="x")
        fila = ttk.Frame(conf)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(leer_ajustes(gen_onedrive.ARCHIVO_AJUSTES)["filas"]))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas,
                    command=self.guardar_filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_filas())

        ttk.Label(conf, text="Profesores con filas reservadas (cada uno recibe su propia contraseña):").pack(anchor="w", pady=(10, 2))
        tabla = ttk.Frame(conf)
        tabla.pack(fill="x")
        self.tabla_prof = ttk.Treeview(tabla, columns=("programa", "profesor", "filas"), show="headings", height=3)
        for col, titulo, ancho in (("programa", "Programa", 140), ("profesor", "Profesor", 420), ("filas", "Filas", 70)):
            self.tabla_prof.heading(col, text=titulo)
            self.tabla_prof.column(col, width=ancho, anchor="w" if col != "filas" else "center")
        self.tabla_prof.pack(side="left", fill="x", expand=True)
        barra = ttk.Scrollbar(tabla, orient="vertical", command=self.tabla_prof.yview)
        barra.pack(side="left", fill="y")
        self.tabla_prof.configure(yscrollcommand=barra.set)

        form = ttk.Frame(conf)
        form.pack(fill="x", pady=(6, 0))
        self.prog = tk.StringVar()
        self.prof = tk.StringVar()
        self.cant = tk.StringVar(value="5")
        ttk.Combobox(form, textvariable=self.prog, values=app.programas, state="readonly", width=14).pack(side="left")
        ttk.Entry(form, textvariable=self.prof, width=26).pack(side="left", padx=6)
        ttk.Spinbox(form, from_=1, to=FILAS_MAX, width=5, textvariable=self.cant).pack(side="left")
        ttk.Button(form, text="Agregar", command=self.agregar_profesor).pack(side="left", padx=6)
        ttk.Button(form, text="Quitar seleccionado", command=self.quitar_profesor).pack(side="left")
        nota(conf, "Programa · nombre · filas. Sin reservas, todo el programa usa la contraseña del programa. "
                   "Los cambios se guardan solos.").pack(anchor="w", pady=(4, 0))

        # 2. Contraseñas
        cont = ttk.LabelFrame(cuerpo, text=" 2. Contraseñas ", padding=8)
        cont.pack(fill="x", pady=(8, 0))
        tabla = ttk.Frame(cont)
        tabla.pack(fill="x")
        self.tabla_pw = ttk.Treeview(tabla, columns=("rol", "ambito", "contrasena"), show="headings", height=4)
        for col, titulo, ancho in (("rol", "Rol", 140), ("ambito", "Para", 330), ("contrasena", "Contraseña", 160)):
            self.tabla_pw.heading(col, text=titulo)
            self.tabla_pw.column(col, width=ancho, anchor="w")
        self.tabla_pw.pack(side="left", fill="x", expand=True)
        barra = ttk.Scrollbar(tabla, orient="vertical", command=self.tabla_pw.yview)
        barra.pack(side="left", fill="y")
        self.tabla_pw.configure(yscrollcommand=barra.set)
        botones = ttk.Frame(cont)
        botones.pack(fill="x", pady=(6, 0))
        self.btn_mostrar = ttk.Button(botones, text="Mostrar contraseñas", command=self.alternar_mostrar)
        self.btn_mostrar.pack(side="left")
        ttk.Button(botones, text="Copiar seleccionada", command=self.copiar_contrasena).pack(side="left", padx=6)
        ttk.Button(botones, text="Regenerar seleccionada", command=lambda: self.regenerar(todas=False)).pack(side="left")
        ttk.Button(botones, text="Regenerar todas", command=lambda: self.regenerar(todas=True)).pack(side="left", padx=6)
        nota(cont, "Archivo: onedrive/privado/contrasenas.csv (no se sube a git). Entregue a cada persona solo la suya.").pack(anchor="w", pady=(4, 0))

        # 3. Generar
        gen = ttk.LabelFrame(cuerpo, text=" 3. Generar ", padding=8)
        gen.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(gen)
        fila.pack(fill="x")
        self.boton = BotonColor(fila, "Generar Excel para OneDrive", COLOR_ONEDRIVE, self.generar)
        self.boton.pack(side="left")
        ttk.Button(fila, text="Abrir carpeta de salida", command=lambda: abrir(gen_onedrive.SALIDA_DIR)).pack(side="left", padx=10)
        ttk.Button(fila, text="Ver guía de despliegue", command=lambda: abrir(ONEDRIVE / "README.md")).pack(side="left")
        nota(gen, "El Excel sale vacío: genérelo antes de empezar a registrar datos y súbalo a OneDrive.").pack(anchor="w", pady=(6, 0))

        self.registro = Registro(cuerpo, "Resultado (OneDrive):")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))

        self.cargar_profesores()
        self.cargar_contrasenas()

    # --- configuración
    def guardar_filas(self):
        try:
            n = int(self.filas.get())
        except ValueError:
            return
        if FILAS_MIN <= n <= FILAS_MAX:
            ajustes = leer_ajustes(gen_onedrive.ARCHIVO_AJUSTES)
            if ajustes["filas"] != n:
                ajustes["filas"] = n
                guardar_ajustes(gen_onedrive.ARCHIVO_AJUSTES, ajustes)

    def cargar_profesores(self):
        self.tabla_prof.delete(*self.tabla_prof.get_children())
        for fila in leer_profesores():
            self.tabla_prof.insert("", "end", values=fila)

    def filas_profesores(self):
        return [tuple(str(v) for v in self.tabla_prof.item(i, "values")) for i in self.tabla_prof.get_children()]

    def agregar_profesor(self):
        programa, profesor, cantidad = self.prog.get(), self.prof.get().strip(), self.cant.get().strip()
        if not programa or not profesor:
            messagebox.showwarning("Falta información", "Elija el programa y escriba el nombre del profesor.", parent=self)
            return
        if not cantidad.isdigit() or int(cantidad) < 1:
            messagebox.showwarning("Filas inválidas", "La cantidad de filas debe ser un número mayor que 0.", parent=self)
            return
        actuales = self.filas_profesores()
        if any(p == programa and n.lower() == profesor.lower() for p, n, _ in actuales):
            messagebox.showwarning("Repetido", f"{profesor} ya tiene filas en {programa}.", parent=self)
            return
        usadas = sum(int(c) for p, _, c in actuales if p == programa) + int(cantidad)
        if usadas > int(self.filas.get() or 0):
            messagebox.showwarning("Sin espacio", f"{programa} solo tiene {self.filas.get()} filas; con esta "
                                   f"reserva serían {usadas}. Aumente las filas por programa.", parent=self)
            return
        self.tabla_prof.insert("", "end", values=(programa, profesor, cantidad))
        guardar_profesores(self.filas_profesores())
        self.prof.set("")

    def quitar_profesor(self):
        seleccion = self.tabla_prof.selection()
        if not seleccion:
            return
        for i in seleccion:
            self.tabla_prof.delete(i)
        guardar_profesores(self.filas_profesores())

    # --- contraseñas
    def cargar_contrasenas(self):
        self.tabla_pw.delete(*self.tabla_pw.get_children())
        for (rol, ambito), clave in gen_onedrive.leer_contrasenas().items():
            visible = clave if self.mostrar else "•" * 10
            self.tabla_pw.insert("", "end", iid=f"{rol}\x1f{ambito}",
                                 values=(ROLES_CONTRASENA.get(rol, rol), ambito, visible))

    def alternar_mostrar(self):
        self.mostrar = not self.mostrar
        self.btn_mostrar.configure(text="Ocultar contraseñas" if self.mostrar else "Mostrar contraseñas")
        self.cargar_contrasenas()

    def clave_seleccionada(self):
        seleccion = self.tabla_pw.selection()
        if not seleccion:
            messagebox.showinfo("Seleccione una fila", "Seleccione una contraseña de la tabla.", parent=self)
            return None
        return tuple(seleccion[0].split("\x1f", 1))

    def copiar_contrasena(self):
        clave = self.clave_seleccionada()
        if clave:
            self.app.copiar(gen_onedrive.leer_contrasenas()[clave])
            self.registro.escribir(f"Contraseña de «{clave[1]}» copiada al portapapeles.")

    def regenerar(self, todas):
        contrasenas = gen_onedrive.leer_contrasenas()
        if not contrasenas:
            messagebox.showinfo("Sin contraseñas", "Primero genere el Excel; ahí se crean las contraseñas.", parent=self)
            return
        claves = list(contrasenas) if todas else [self.clave_seleccionada()]
        if None in claves:
            return
        que = "TODAS las contraseñas" if todas else f"la contraseña de «{claves[0][1]}»"
        if not messagebox.askyesno(
                "Regenerar contraseñas",
                f"Se cambiará {que} y se volverá a generar el Excel (vacío).\n\n"
                "Si ya hay un Excel en uso en OneDrive, tendrá que reemplazarlo por el nuevo y "
                "entregar las contraseñas nuevas. ¿Continuar?", parent=self):
            return
        for clave in claves:
            contrasenas[clave] = gen_onedrive.generar_contrasena()
        gen_onedrive.guardar_contrasenas(contrasenas)
        self.registro.escribir(f"Se regeneró {que}.", limpiar=True)
        self.generar(limpiar=False)

    # --- generar
    def generar(self, limpiar=True):
        self.guardar_filas()
        guardar_profesores(self.filas_profesores())
        self.app.ejecutar(gen_onedrive.__file__, self.registro, "OneDrive", limpiar, self.cargar_contrasenas)


# ---------------------------------------------------------------------------
# Pestaña Google Sheets
# ---------------------------------------------------------------------------

class PestanaGoogle(ttk.Frame):
    def __init__(self, app, padre):
        super().__init__(padre, padding=(0, 0, 0, 8))
        self.app = app

        encabezado(self, COLOR_GOOGLE, "Plantilla para Google Sheets",
                   "Permisos por cuenta institucional.   Entradas: raw/ + google-sheets/config/   ·   "
                   "Salida: google-sheets/salida/ + Apps Script").pack(fill="x")
        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)

        # 1. Plantilla
        conf = ttk.LabelFrame(cuerpo, text=" 1. Configuración de la plantilla ", padding=8)
        conf.pack(fill="x")
        fila = ttk.Frame(conf)
        fila.pack(fill="x")
        ttk.Label(fila, text="Filas de registro por programa:").pack(side="left")
        self.filas = tk.StringVar(value=str(leer_ajustes(gen_google.ARCHIVO_AJUSTES)["filas"]))
        ttk.Spinbox(fila, from_=FILAS_MIN, to=FILAS_MAX, increment=10, width=7, textvariable=self.filas,
                    command=self.guardar_filas).pack(side="left", padx=6)
        self.filas.trace_add("write", lambda *_: self.guardar_filas())
        nota(conf, "Los usuarios y sus roles no se configuran aquí: se administran en la hoja «Usuarios» "
                   "dentro de Google Sheets.").pack(anchor="w", pady=(6, 0))

        # 2. Apps Script
        script = ttk.LabelFrame(cuerpo, text=" 2. Configuración del Apps Script ", padding=8)
        script.pack(fill="x", pady=(8, 0))
        valores = leer_config_script()
        fila = ttk.Frame(script)
        fila.pack(fill="x")
        ttk.Label(fila, text="Dominio institucional:  @").pack(side="left")
        self.dominio = tk.StringVar(value=valores["DOMINIO"])
        entrada = ttk.Entry(fila, textvariable=self.dominio, width=24)
        entrada.pack(side="left")
        entrada.bind("<FocusOut>", lambda _e: self.guardar_script())
        entrada.bind("<Return>", lambda _e: self.guardar_script())
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
        self.estado_script = nota(script, "")
        self.estado_script.pack(anchor="w", pady=(4, 0))
        nota(script, "Si el script ya está instalado en Google, después de cambiar esto vuelva a copiar Codigo.gs "
                     "y péguelo en Extensiones → Apps Script.").pack(anchor="w")

        # 3. Generar
        gen = ttk.LabelFrame(cuerpo, text=" 3. Generar y copiar ", padding=8)
        gen.pack(fill="x", pady=(8, 0))
        fila = ttk.Frame(gen)
        fila.pack(fill="x")
        self.boton = BotonColor(fila, "Generar plantilla para Google Sheets", COLOR_GOOGLE, self.generar)
        self.boton.pack(side="left")
        ttk.Button(fila, text="Abrir carpeta de salida", command=lambda: abrir(gen_google.SALIDA_DIR)).pack(side="left", padx=10)
        ttk.Button(fila, text="Ver guía de despliegue", command=lambda: abrir(GOOGLE / "README.md")).pack(side="left")
        fila = ttk.Frame(gen)
        fila.pack(fill="x", pady=(8, 0))
        ttk.Button(fila, text="Copiar Codigo.gs", command=lambda: self.copiar_archivo(CODIGO_GS)).pack(side="left")
        ttk.Button(fila, text="Copiar appsscript.json", command=lambda: self.copiar_archivo(MANIFIESTO)).pack(side="left", padx=6)
        nota(gen, "Importe la plantilla a Google Sheets y pegue el código en Extensiones → Apps Script "
                  "(pasos completos en la guía).").pack(anchor="w", pady=(6, 0))

        self.registro = Registro(cuerpo, "Resultado (Google Sheets):")
        self.registro.pack(fill="both", expand=True, pady=(8, 0))

    def guardar_filas(self):
        try:
            n = int(self.filas.get())
        except ValueError:
            return
        if FILAS_MIN <= n <= FILAS_MAX:
            ajustes = leer_ajustes(gen_google.ARCHIVO_AJUSTES)
            if ajustes["filas"] != n:
                ajustes["filas"] = n
                guardar_ajustes(gen_google.ARCHIVO_AJUSTES, ajustes)

    def guardar_script(self):
        dominio = self.dominio.get().strip().lstrip("@").lower()
        if not re.fullmatch(r"[a-z0-9-]+(\.[a-z0-9-]+)+", dominio):
            self.estado_script.configure(text="⚠️ Dominio inválido (ejemplo: uniandes.edu.co). No se guardó.", foreground="#C5221F")
            return False
        try:
            hora = int(self.hora.get())
            assert 0 <= hora <= 23
        except (ValueError, AssertionError):
            self.estado_script.configure(text="⚠️ La hora debe estar entre 0 y 23. No se guardó.", foreground="#C5221F")
            return False
        nuevos = {"DOMINIO": dominio, "HORA_RESTAURACION": hora,
                  "GESTOR_EDITA_FILAS_DE_PROFESORES": self.gestores.get(), "ALERTAS_POR_CORREO": self.alertas.get()}
        if nuevos != leer_config_script():
            guardar_config_script(nuevos)
            self.estado_script.configure(text="✓ Cambios guardados en google-sheets/apps-script/Codigo.gs", foreground=COLOR_GOOGLE)
        return True

    def copiar_archivo(self, ruta):
        if ruta == CODIGO_GS and not self.guardar_script():
            return
        self.app.copiar(ruta.read_text(encoding="utf-8"))
        self.registro.escribir(f"{ruta.name} copiado al portapapeles. Péguelo en Extensiones → Apps Script.")

    def generar(self):
        self.guardar_filas()
        self.app.ejecutar(gen_google.__file__, self.registro, "Google Sheets", True)


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cronograma 2027 Bloque Neón — Generador")
        self.ocupado = False

        plantillas, _ = leer_plantillas()
        self.programas = list(plantillas)

        arriba = ttk.Frame(self, padding=(12, 10, 12, 6))
        arriba.pack(fill="x")
        ttk.Label(arriba, text=f"Entrada común para las dos versiones: raw/  ({len(self.programas)} programas)").pack(side="left")
        ttk.Button(arriba, text="Abrir raw/", command=lambda: abrir(RAW_DIR)).pack(side="right")

        self.pestanas = ttk.Notebook(self)
        self.pestanas.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.onedrive = PestanaOneDrive(self, self.pestanas)
        self.google = PestanaGoogle(self, self.pestanas)
        self.pestanas.add(self.onedrive, text="   Excel para OneDrive   ")
        self.pestanas.add(self.google, text="   Google Sheets   ")

        # Tamaño según el contenido, sin pasarse de la pantalla.
        self.update_idletasks()
        alto = min(self.winfo_reqheight(), self.winfo_screenheight() - 80)
        self.geometry(f"{max(self.winfo_reqwidth(), 900)}x{alto}")
        self.minsize(860, min(alto, 640))

    def copiar(self, texto):
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.update()

    def ejecutar(self, script, registro, nombre, limpiar, al_terminar=None):
        """Corre un generador en segundo plano y muestra su salida en el registro de su pestaña."""
        if self.ocupado:
            return
        self.ocupado = True
        for pestana in (self.onedrive, self.google):
            pestana.boton.habilitar(False)
        registro.escribir(f"Generando versión {nombre}…", limpiar=limpiar)
        # Tkinter solo se puede tocar desde el hilo principal: el hilo de trabajo deja el
        # resultado en una cola y la ventana la revisa cada 100 ms.
        resultado = queue.Queue()

        def trabajo():
            resultado.put(subprocess.run([sys.executable, str(script)], cwd=ROOT, capture_output=True, text=True))

        def revisar():
            try:
                proceso = resultado.get_nowait()
            except queue.Empty:
                self.after(100, revisar)
                return
            terminar(proceso)

        def terminar(proceso):
            salida = (proceso.stdout + proceso.stderr).strip()
            registro.escribir(salida or "(sin salida)")
            registro.escribir("✓ Listo." if proceso.returncode == 0 else "✗ No se pudo generar. Revise el mensaje de arriba.")
            self.ocupado = False
            for pestana in (self.onedrive, self.google):
                pestana.boton.habilitar(True)
            if al_terminar:
                al_terminar()

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(100, revisar)


if __name__ == "__main__":
    App().mainloop()
