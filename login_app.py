""""
Inicio de sesión gráfico
"""
import hashlib
import hmac
import os
import re
import sqlite3
import sys
import tkinter as tk
from contextlib import closing
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from openpyxl.styles import Font

# La base de datos se guarda junto al programa (.py o .exe), sin importar desde dónde se abra
BASE = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                       else os.path.abspath(__file__))
DB = os.path.join(BASE, "usuarios.db")

# Dominios de correo aceptados (edita esta lista si necesitas otros)
DOMINIOS_CORREO = (
    "gmail.com",
    "outlook.com",
    "hotmail.com",
    "live.com",
    "yahoo.com",
    "icloud.com",
    "comunidad.unam.mx",
    "aragon.unam.mx",
)

# Requisitos de contraseña: (texto que se muestra, función que lo verifica)
REQUISITOS_PASSWORD = (
    ("Mínimo 8 caracteres", lambda p: len(p) >= 8),
    ("Al menos una letra mayúscula (A-Z)", lambda p: any(c.isupper() for c in p)),
    ("Al menos una letra minúscula (a-z)", lambda p: any(c.islower() for c in p)),
    ("Al menos un número (0-9)", lambda p: any(c.isdigit() for c in p)),
    ("Al menos un carácter especial (! @ # $ % & * ...)",
     lambda p: any(not c.isalnum() and not c.isspace() for c in p)),
    ("Sin espacios", lambda p: not any(c.isspace() for c in p)),
)

VERDE, ROJO, GRIS = "#2e7d32", "#b00020", "#666666"
PATRON_CORREO = re.compile(r"^[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+)$")


# ---------------------------- Base de datos ----------------------------
def conectar():
    return sqlite3.connect(DB)


def init_db():
    with closing(conectar()) as c, c:
        c.execute(
            """CREATE TABLE IF NOT EXISTS usuarios (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT UNIQUE NOT NULL,
                correo  TEXT NOT NULL,
                salt    BLOB NOT NULL,
                hash    BLOB NOT NULL,
                creado  TEXT NOT NULL
            )"""
        )


def hashear(password, salt):
    return hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 200_000)


def registrar(usuario, correo, password):
    salt = os.urandom(16)
    try:
        with closing(conectar()) as c, c:
            c.execute(
                "INSERT INTO usuarios (usuario, correo, salt, hash, creado) VALUES (?,?,?,?,?)",
                (usuario, correo, salt, hashear(password, salt),
                 datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )
        return True
    except sqlite3.IntegrityError:
        return False  # el usuario ya existe


def validar(usuario, password):
    with closing(conectar()) as c:
        fila = c.execute(
            "SELECT salt, hash FROM usuarios WHERE usuario = ?", (usuario,)
        ).fetchone()
    return bool(fila) and hmac.compare_digest(hashear(password, fila[0]), fila[1])


def exportar_excel(ruta):
    """Exporta los usuarios a un .xlsx (sin salt ni hash de las contraseñas)."""
    with closing(conectar()) as c:
        filas = c.execute(
            "SELECT id, usuario, correo, creado FROM usuarios ORDER BY id"
        ).fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "Usuarios"
    ws.append(["ID", "Usuario", "Correo", "Fecha de registro"])
    for celda in ws[1]:
        celda.font = Font(bold=True)
    for fila in filas:
        ws.append(fila)
    for col, ancho in zip("ABCD", (6, 20, 32, 22)):
        ws.column_dimensions[col].width = ancho
    wb.save(ruta)


# ------------------------------ Validaciones ------------------------------
def correo_valido(correo):
    """Devuelve (True, '') o (False, motivo)."""
    m = PATRON_CORREO.match(correo)
    if not m:
        return False, "El formato del correo no es válido (ej. nombre@gmail.com)."
    if m.group(1).lower() not in DOMINIOS_CORREO:
        lista = ", ".join("@" + d for d in DOMINIOS_CORREO)
        return False, f"Ese dominio no está permitido.\n\nDominios aceptados:\n{lista}"
    return True, ""


def requisitos_faltantes(password):
    return [texto for texto, ok in REQUISITOS_PASSWORD if not ok(password)]


# ------------------- Contenedor con barras de desplazamiento -------------------
class MarcoDesplazable(ttk.Frame):
    """Frame que muestra barras de desplazamiento solo cuando el contenido no cabe."""

    BARRA = 17  # grosor aproximado de una barra

    def __init__(self, padre):
        super().__init__(padre)
        fondo = ttk.Style().lookup("TFrame", "background") or "#f0f0f0"
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=fondo, yscrollincrement=20)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.interior = ttk.Frame(self.canvas)
        self.interior.columnconfigure(0, weight=1)
        self.interior.rowconfigure(0, weight=1)
        self.ventana = self.canvas.create_window((0, 0), window=self.interior, anchor="nw")

        self._v = self._h = False
        self.bind("<Configure>", self._ajustar)
        self.interior.bind("<Configure>", self._ajustar)

        # Rueda del mouse y navegación con Tab
        self.bind_all("<MouseWheel>", self._rueda)   # Windows / macOS
        self.bind_all("<Button-4>", self._rueda)     # Linux (arriba)
        self.bind_all("<Button-5>", self._rueda)     # Linux (abajo)
        self.bind_all("<FocusIn>", self._mostrar_foco)

    def _ajustar(self, _evento=None):
        tw, th = self.winfo_width(), self.winfo_height()
        rw, rh = self.interior.winfo_reqwidth(), self.interior.winfo_reqheight()
        if tw <= 1 or th <= 1:
            return
        v = rh > th
        h = rw > tw - (self.BARRA if v else 0)
        v = v or (h and rh > th - self.BARRA)

        if v != self._v:
            self._v = v
            self.vsb.grid(row=0, column=1, sticky="ns") if v else self.vsb.grid_remove()
        if h != self._h:
            self._h = h
            self.hsb.grid(row=1, column=0, sticky="ew") if h else self.hsb.grid_remove()

        ancho = max(tw - (self.BARRA if v else 0), rw)
        alto = max(th - (self.BARRA if h else 0), rh)
        self.canvas.itemconfigure(self.ventana, width=ancho, height=alto)
        self.canvas.configure(scrollregion=(0, 0, ancho, alto))
        if not v:
            self.canvas.yview_moveto(0)
        if not h:
            self.canvas.xview_moveto(0)

    def _rueda(self, e):
        if not self._v:
            return
        if getattr(e, "num", None) == 4:
            paso = -1
        elif getattr(e, "num", None) == 5:
            paso = 1
        else:
            paso = -1 if e.delta > 0 else 1
        self.canvas.yview_scroll(paso, "units")

    def _mostrar_foco(self, e):
        """Si un campo queda fuera de la vista al navegar con Tab, desplaza hasta él."""
        try:
            w = e.widget
            if not self._v or not str(w).startswith(str(self.interior)):
                return
            y = w.winfo_rooty() - self.interior.winfo_rooty()
            alto_vista = self.canvas.winfo_height()
            total = max(self.interior.winfo_height(), 1)
            arriba = self.canvas.canvasy(0)
            if y < arriba:
                self.canvas.yview_moveto(max(y - 10, 0) / total)
            elif y + w.winfo_height() > arriba + alto_vista:
                self.canvas.yview_moveto((y + w.winfo_height() + 10 - alto_vista) / total)
        except tk.TclError:
            pass

    def destroy(self):
        for evento in ("<MouseWheel>", "<Button-4>", "<Button-5>", "<FocusIn>"):
            self.unbind_all(evento)
        super().destroy()


# ------------------------------ Interfaz ------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de acceso")
        self.geometry("540x660")
        self.minsize(300, 240)          # la ventana se puede redimensionar
        self.contenedor = ttk.Frame(self)
        self.contenedor.pack(fill="both", expand=True)
        self.mostrar_login()

    def limpiar(self):
        for w in self.contenedor.winfo_children():
            w.destroy()

    def pantalla_formulario(self):
        """Crea una pantalla con desplazamiento y devuelve el frame donde van los campos."""
        self.limpiar()
        marco = MarcoDesplazable(self.contenedor)
        marco.pack(fill="both", expand=True)
        form = ttk.Frame(marco.interior, padding=20)
        form.grid(row=0, column=0)      # centrado en la ventana
        return form

    @staticmethod
    def campo(padre, texto, oculto=False):
        ttk.Label(padre, text=texto).pack(anchor="w", pady=(10, 0))
        e = ttk.Entry(padre, show="*" if oculto else "", width=38)
        e.pack(fill="x")
        return e

    # --- Pantalla de login ---
    def mostrar_login(self):
        form = self.pantalla_formulario()
        ttk.Label(form, text="Iniciar sesión", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))
        usuario = self.campo(form, "Usuario")
        password = self.campo(form, "Contraseña", oculto=True)

        def entrar(_evento=None):
            u, p = usuario.get().strip(), password.get()
            if validar(u, p):
                self.mostrar_panel(u)
            else:
                messagebox.showerror("Error", "Usuario o contraseña incorrectos")

        password.bind("<Return>", entrar)
        ttk.Button(form, text="Entrar", command=entrar).pack(pady=(18, 6), fill="x")
        ttk.Button(form, text="Crear cuenta", command=self.mostrar_registro).pack(fill="x")
        usuario.focus()

    # --- Pantalla de registro ---
    def mostrar_registro(self):
        form = self.pantalla_formulario()
        ttk.Label(form, text="Crear cuenta", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))

        usuario = self.campo(form, "Usuario")

        correo = self.campo(form, "Correo")
        dominios = ", ".join("@" + d for d in DOMINIOS_CORREO)
        ttk.Label(form, text=f"Correos aceptados: {dominios}", foreground=GRIS,
                  wraplength=330, justify="left").pack(anchor="w", pady=(2, 0))

        password = self.campo(form, "Contraseña", oculto=True)
        ttk.Label(form, text="La contraseña debe cumplir:", foreground=GRIS).pack(anchor="w", pady=(4, 0))
        etiquetas = []
        for texto, _ in REQUISITOS_PASSWORD:
            lbl = ttk.Label(form, text=f"•  {texto}", foreground=GRIS,
                            wraplength=330, justify="left")
            lbl.pack(anchor="w")
            etiquetas.append(lbl)

        confirmar = self.campo(form, "Confirmar contraseña", oculto=True)
        coincide = ttk.Label(form, text="", foreground=GRIS)
        coincide.pack(anchor="w", pady=(2, 0))

        def actualizar(_evento=None):
            p = password.get()
            for (texto, ok), lbl in zip(REQUISITOS_PASSWORD, etiquetas):
                if not p:
                    lbl.config(text=f"•  {texto}", foreground=GRIS)
                elif ok(p):
                    lbl.config(text=f"✔  {texto}", foreground=VERDE)
                else:
                    lbl.config(text=f"✘  {texto}", foreground=ROJO)
            c = confirmar.get()
            if not c:
                coincide.config(text="", foreground=GRIS)
            elif c == p:
                coincide.config(text="✔  Las contraseñas coinciden", foreground=VERDE)
            else:
                coincide.config(text="✘  Las contraseñas no coinciden", foreground=ROJO)

        password.bind("<KeyRelease>", actualizar)
        confirmar.bind("<KeyRelease>", actualizar)

        def guardar():
            u = usuario.get().strip()
            c = correo.get().strip().lower()
            p = password.get()
            if not u or not c or not p:
                return messagebox.showwarning("Faltan datos", "Completa todos los campos")
            ok, motivo = correo_valido(c)
            if not ok:
                return messagebox.showwarning("Correo no válido", motivo)
            faltan = requisitos_faltantes(p)
            if faltan:
                detalle = "\n".join("• " + t for t in faltan)
                return messagebox.showwarning("Contraseña insegura",
                                              f"A tu contraseña le falta:\n\n{detalle}")
            if p != confirmar.get():
                return messagebox.showwarning("Contraseña", "Las contraseñas no coinciden")
            if registrar(u, c, p):
                messagebox.showinfo("Listo", "Cuenta creada. Ya puedes iniciar sesión.")
                self.mostrar_login()
            else:
                messagebox.showerror("Error", "Ese usuario ya existe")

        ttk.Button(form, text="Registrarme", command=guardar).pack(pady=(14, 6), fill="x")
        ttk.Button(form, text="Volver", command=self.mostrar_login).pack(fill="x")
        usuario.focus()

    # --- Pantalla tras iniciar sesión ---
    def mostrar_panel(self, usuario):
        self.limpiar()
        marco = ttk.Frame(self.contenedor, padding=20)
        marco.pack(fill="both", expand=True)
        centro = ttk.Frame(marco)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(centro, text="✔", font=("Segoe UI", 40), foreground=VERDE).pack()
        ttk.Label(centro, text="Has iniciado sesión correctamente",
                  font=("Segoe UI", 14, "bold"), wraplength=320,
                  justify="center").pack(pady=(6, 18))
        def exportar():
            ruta = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                initialfile="usuarios.xlsx")
            if not ruta:
                return
            try:
                exportar_excel(ruta)
                messagebox.showinfo("Exportado", f"Archivo guardado en:\n{ruta}")
            except PermissionError:
                messagebox.showerror("Error", "Cierra el archivo si lo tienes abierto en Excel")

        ttk.Button(centro, text="Exportar a Excel", command=exportar).pack(pady=(0, 8))
        ttk.Button(centro, text="Cerrar sesión", command=self.mostrar_login).pack()


if __name__ == "__main__":
    init_db()
    App().mainloop()
