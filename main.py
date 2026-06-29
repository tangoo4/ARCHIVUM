import customtkinter as ctk
from tkinter import filedialog, messagebox


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

APP_NOMBRE = "ARCHIVUM"
APP_VERSION = "v0.1"

ANCHO = 1200
ALTO = 750

COLOR_FONDO = "#121212"
COLOR_PANEL = "#1E1E1E"
COLOR_PANEL_2 = "#2A2A2A"
COLOR_TEXTO = "#FFFFFF"
COLOR_TEXTO_2 = "#B8B8B8"
COLOR_VERDE = "#00FF66"
COLOR_VERDE_HOVER = "#00CC55"
COLOR_CIAN = "#00E5FF"
COLOR_ROJO = "#FF3030"

FUENTE_TITULO = ("Segoe UI", 32, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 20, "bold")
FUENTE_NORMAL = ("Segoe UI", 15)
FUENTE_BOTON = ("Segoe UI", 18, "bold")


# ==========================================================
# APP PRINCIPAL
# ==========================================================

class ArchivumApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NOMBRE)
        self.geometry("1200x750")
        self.state("zoomed")
        self.resizable(True, True)
        self.configure(fg_color=COLOR_FONDO)

        self.pantalla_actual = None
        self.archivo_actual = None

        self.mostrar_inicio()

    def limpiar_pantalla(self):
        if self.pantalla_actual is not None:
            self.pantalla_actual.destroy()
            self.pantalla_actual = None

    def mostrar_inicio(self):
        self.limpiar_pantalla()
        self.pantalla_actual = PantallaInicio(
            self,
            abrir_nueva=self.mostrar_nueva_temporada,
            abrir_medido=self.mostrar_medido
        )

    def mostrar_nueva_temporada(self):
        self.limpiar_pantalla()
        self.pantalla_actual = PantallaNuevaTemporada(
            self,
            volver=self.mostrar_inicio,
            abrir_medido=self.mostrar_medido
        )

    def mostrar_medido(self):
        self.limpiar_pantalla()
        self.pantalla_actual = PantallaMedido(
            self,
            volver=self.mostrar_inicio
        )


# ==========================================================
# COMPONENTES SIMPLES
# ==========================================================

def boton_principal(master, texto, comando):
    return ctk.CTkButton(
        master,
        text=texto,
        command=comando,
        width=340,
        height=60,
        fg_color=COLOR_VERDE,
        hover_color=COLOR_VERDE_HOVER,
        text_color="#000000",
        font=FUENTE_BOTON,
        corner_radius=12
    )


def boton_secundario(master, texto, comando):
    return ctk.CTkButton(
        master,
        text=texto,
        command=comando,
        width=220,
        height=45,
        fg_color=COLOR_PANEL_2,
        hover_color="#3A3A3A",
        text_color=COLOR_TEXTO,
        font=FUENTE_NORMAL,
        corner_radius=10
    )


def entrada(master, ancho=360):
    campo = ctk.CTkEntry(
        master,
        width=ancho,
        height=42,
        fg_color=COLOR_PANEL_2,
        border_color="#444444",
        text_color=COLOR_TEXTO,
        font=FUENTE_NORMAL,
        corner_radius=8
    )
    return campo


# ==========================================================
# PANTALLA INICIO
# ==========================================================

class PantallaInicio(ctk.CTkFrame):
    def __init__(self, master, abrir_nueva, abrir_medido):
        super().__init__(master, fg_color=COLOR_FONDO)
        self.master = master
        self.abrir_nueva = abrir_nueva
        self.abrir_medido = abrir_medido

        self.pack(fill="both", expand=True)
        self.crear()

    def crear(self):
        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PANEL, height=100, corner_radius=0)
        cabecera.pack(fill="x")

        ctk.CTkLabel(
            cabecera,
            text="ARCHIVUM",
            font=FUENTE_TITULO,
            text_color=COLOR_VERDE
        ).place(x=40, y=20)

        ctk.CTkLabel(
            cabecera,
            text="Sistema de Gestión de Encuadernación Notarial",
            font=FUENTE_NORMAL,
            text_color=COLOR_TEXTO_2
        ).place(x=42, y=65)

        ctk.CTkLabel(
            cabecera,
            text=APP_VERSION,
            font=FUENTE_NORMAL,
            text_color=COLOR_CIAN
        ).place(x=1450, y=40)

        panel = ctk.CTkFrame(
            self,
            width=640,
            height=460,
            fg_color=COLOR_PANEL,
            corner_radius=16,
            border_width=1,
            border_color="#3A3A3A"
        )
        panel.place(relx=0.5, rely=0.52, anchor="center")

        ctk.CTkLabel(
            panel,
            text="SELECCIONAR MODO DE TRABAJO",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).pack(pady=(50, 40))

        boton_principal(
            panel,
            "CONTINUAR TEMPORADA",
            self.continuar_temporada
        ).pack(pady=16)

        boton_principal(
            panel,
            "NUEVA TEMPORADA",
            self.abrir_nueva
        ).pack(pady=16)

        boton_principal(
            panel,
            "ÚLTIMO ARCHIVO",
            self.ultimo_archivo
        ).pack(pady=16)

        self.estado = ctk.CTkLabel(
            self,
            text="Estado: esperando selección...",
            font=FUENTE_NORMAL,
            text_color=COLOR_TEXTO_2
        )
        self.estado.place(x=30, y=850)

    def continuar_temporada(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )

        if not archivo:
            self.estado.configure(text="Estado: selección cancelada")
            return

        self.master.archivo_actual = archivo
        self.estado.configure(text=f"Estado: archivo seleccionado: {archivo}", text_color=COLOR_VERDE)

        messagebox.showinfo(
            "Archivo seleccionado",
            f"Archivo cargado:\n\n{archivo}\n\nEn v0.1 todavía no se lee Excel."
        )

        self.abrir_medido()

    def ultimo_archivo(self):
        messagebox.showinfo("Pendiente", "Último archivo se añadirá en una versión posterior.")


# ==========================================================
# PANTALLA NUEVA TEMPORADA
# ==========================================================

class PantallaNuevaTemporada(ctk.CTkFrame):
    def __init__(self, master, volver, abrir_medido):
        super().__init__(master, fg_color=COLOR_FONDO)
        self.master = master
        self.volver = volver
        self.abrir_medido = abrir_medido

        self.pack(fill="both", expand=True)
        self.crear()

    def crear(self):
        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PANEL, height=100, corner_radius=0)
        cabecera.pack(fill="x")

        ctk.CTkLabel(
            cabecera,
            text="NUEVA TEMPORADA",
            font=FUENTE_TITULO,
            text_color=COLOR_VERDE
        ).place(x=40, y=20)

        ctk.CTkLabel(
            cabecera,
            text="Creación visual de temporada. En v0.1 todavía no crea Excel.",
            font=FUENTE_NORMAL,
            text_color=COLOR_TEXTO_2
        ).place(x=42, y=65)

        panel = ctk.CTkFrame(
            self,
            width=720,
            height=600,
            fg_color=COLOR_PANEL,
            corner_radius=16,
            border_width=1,
            border_color="#3A3A3A"
        )
        panel.place(relx=0.5, rely=0.54, anchor="center")

        ctk.CTkLabel(
            panel,
            text="DATOS DE TEMPORADA",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).pack(pady=(35, 25))

        ctk.CTkLabel(panel, text="TIPO", font=FUENTE_NORMAL, text_color=COLOR_TEXTO).pack(anchor="w", padx=175)
        self.tipo = ctk.CTkComboBox(
            panel,
            values=["PROTOCOLO", "POLIZAS", "LIBRO INDICADOR"],
            width=360,
            height=42,
            fg_color=COLOR_PANEL_2,
            button_color=COLOR_VERDE,
            button_hover_color=COLOR_VERDE_HOVER,
            text_color=COLOR_TEXTO,
            font=FUENTE_NORMAL
        )
        self.tipo.set("PROTOCOLO")
        self.tipo.pack(pady=(6, 18))

        ctk.CTkLabel(panel, text="NOTARIO", font=FUENTE_NORMAL, text_color=COLOR_TEXTO).pack(anchor="w", padx=175)
        self.notario = entrada(panel)
        self.notario.pack(pady=(6, 18))

        ctk.CTkLabel(panel, text="AÑO", font=FUENTE_NORMAL, text_color=COLOR_TEXTO).pack(anchor="w", padx=175)
        self.anio = entrada(panel)
        self.anio.pack(pady=(6, 18))

        ctk.CTkLabel(panel, text="MEDIDA ESTÁNDAR", font=FUENTE_NORMAL, text_color=COLOR_TEXTO).pack(anchor="w", padx=175)
        self.medida = entrada(panel)
        self.medida.pack(pady=(6, 25))

        botones = ctk.CTkFrame(panel, fg_color="transparent")
        botones.pack(pady=10)

        boton_secundario(botones, "VOLVER", self.volver).pack(side="left", padx=12)
        boton_principal(botones, "CREAR TEMPORADA", self.crear_temporada).pack(side="left", padx=12)

        self.notario.focus_set()

    def crear_temporada(self):
        tipo = self.tipo.get()
        notario = self.notario.get().strip().upper()
        anio = self.anio.get().strip()
        medida = self.medida.get().strip()

        if not notario or not anio or not medida:
            messagebox.showwarning("Faltan datos", "Rellena notario, año y medida estándar.")
            return

        self.master.archivo_actual = f"{tipo}_{notario}_{anio}.xlsx"

        messagebox.showinfo(
            "Temporada creada",
            f"Temporada preparada:\n\n{self.master.archivo_actual}\n\nEn v0.1 todavía no se crea Excel."
        )

        self.abrir_medido()


# ==========================================================
# PANTALLA MEDIDO
# ==========================================================

class PantallaMedido(ctk.CTkFrame):
    def __init__(self, master, volver):
        super().__init__(master, fg_color=COLOR_FONDO)
        self.master = master
        self.volver = volver
        self.tomo_actual = 1
        self.medida_estandar = "8,5"

        self.pack(fill="both", expand=True)
        self.crear()

    def crear(self):
        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PANEL, height=100, corner_radius=0)
        cabecera.pack(fill="x")

        archivo = self.master.archivo_actual or "SIN ARCHIVO"

        ctk.CTkLabel(
            cabecera,
            text=f"ARCHIVO: {archivo}",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).place(x=30, y=20)

        ctk.CTkLabel(
            cabecera,
            text="● GUARDADO",
            font=FUENTE_NORMAL,
            text_color=COLOR_VERDE
        ).place(x=30, y=60)

        panel_superior = ctk.CTkFrame(
            self,
            width=1540,
            height=360,
            fg_color=COLOR_PANEL,
            corner_radius=16,
            border_width=1,
            border_color="#3A3A3A"
        )
        panel_superior.place(x=30, y=120)

        ctk.CTkLabel(
            panel_superior,
            text="TOMO ACTUAL",
            font=FUENTE_NORMAL,
            text_color=COLOR_TEXTO_2
        ).place(x=45, y=35)

        self.label_tomo = ctk.CTkLabel(
            panel_superior,
            text=str(self.tomo_actual),
            font=("Segoe UI", 42, "bold"),
            text_color=COLOR_VERDE,
            fg_color=COLOR_PANEL_2,
            width=160,
            height=90,
            corner_radius=14
        )
        self.label_tomo.place(x=45, y=70)

        x_label = 260
        x_entry = 450
        y = 35
        salto = 52

        self.campos = []

        self.crear_campo(panel_superior, "MATRIZ INICIO", x_label, x_entry, y)
        y += salto
        self.crear_campo(panel_superior, "FECHA INICIO", x_label, x_entry, y)
        y += salto
        self.crear_campo(panel_superior, "MATRIZ FINAL", x_label, x_entry, y)
        y += salto
        self.crear_campo(panel_superior, "FECHA FINAL", x_label, x_entry, y)
        y += salto
        self.campo_medida = self.crear_campo(panel_superior, "MEDIDA", x_label, x_entry, y, ancho=150)
        self.campo_medida.insert(0, self.medida_estandar)
        y += salto
        self.crear_campo(panel_superior, "OBSERVACIONES", x_label, x_entry, y, ancho=620)

        for i, campo in enumerate(self.campos):
            campo.bind("<KeyRelease>", self.mayusculas)
            campo.bind("<Return>", lambda event, idx=i: self.enter_campo(idx))

        boton_secundario(panel_superior, "BUSCADOR", self.buscador).place(x=1180, y=105)
        boton_secundario(panel_superior, "BACKUP", self.backup).place(x=1180, y=165)
        boton_secundario(panel_superior, "CERRAR", self.volver).place(x=1180, y=225)

        self.mensaje = ctk.CTkLabel(
            panel_superior,
            text="",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_VERDE
        )
        self.mensaje.place(x=45, y=250)

        panel_inferior = ctk.CTkFrame(
            self,
            width=1540,
            height=360,
            fg_color=COLOR_PANEL,
            corner_radius=16,
            border_width=1,
            border_color="#3A3A3A"
        )
        panel_inferior.place(x=30, y=505)

        ctk.CTkLabel(
            panel_inferior,
            text="ÚLTIMOS TOMOS",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).place(x=25, y=20)

        self.tabla_texto = ctk.CTkTextbox(
            panel_inferior,
            width=1490,
            height=265,
            fg_color="#181818",
            text_color=COLOR_TEXTO,
            font=("Consolas", 16)
        )
        self.tabla_texto.place(x=25, y=70)
        self.tabla_texto.insert("end", "TOMO | INICIO | FINAL | MEDIDA | OBSERVACIONES\n")
        self.tabla_texto.insert("end", "-" * 90 + "\n")
        self.tabla_texto.configure(state="disabled")

        self.campos[0].focus_set()

    def crear_campo(self, master, etiqueta, x_label, x_entry, y, ancho=260):
        ctk.CTkLabel(
            master,
            text=etiqueta,
            font=FUENTE_NORMAL,
            text_color=COLOR_TEXTO
        ).place(x=x_label, y=y)

        campo = entrada(master, ancho=ancho)
        campo.place(x=x_entry, y=y - 5)
        self.campos.append(campo)
        return campo

    def mayusculas(self, event):
        campo = event.widget
        texto = campo.get()
        cursor = campo.index("insert")
        campo.delete(0, "end")
        campo.insert(0, texto.upper())
        campo.icursor(cursor)

    def enter_campo(self, idx):
        if idx < len(self.campos) - 1:
            self.campos[idx + 1].focus_set()
        else:
            self.guardar_tomo()

        return "break"

    def guardar_tomo(self):
        valores = [c.get().strip() for c in self.campos]

        obligatorios = valores[:-1]

        if any(v == "" for v in obligatorios):
            messagebox.showwarning("Faltan datos", "Todos los campos excepto observaciones son obligatorios.")
            return

        linea = f"{self.tomo_actual:<5}| {valores[0]:<7}| {valores[2]:<6}| {valores[4]:<7}| {valores[5]}\n"

        self.tabla_texto.configure(state="normal")
        self.tabla_texto.insert("end", linea)
        self.tabla_texto.see("end")
        self.tabla_texto.configure(state="disabled")

        self.mensaje.configure(text=f"✔ TOMO {self.tomo_actual} GUARDADO")
        self.after(1000, lambda: self.mensaje.configure(text=""))

        self.tomo_actual += 1
        self.label_tomo.configure(text=str(self.tomo_actual))

        for campo in self.campos:
            campo.delete(0, "end")

        self.campo_medida.insert(0, self.medida_estandar)
        self.campos[0].focus_set()

    def buscador(self):
        messagebox.showinfo("Buscador", "Buscador pendiente en v0.2.")

    def backup(self):
        messagebox.showinfo("Backup", "Backup pendiente en v0.2.")


# ==========================================================
# EJECUCIÓN
# ==========================================================

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("green")

    app = ArchivumApp()
    app.mainloop()