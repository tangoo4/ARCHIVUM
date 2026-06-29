"""
Pantalla de inicio de Archivum
Versión 0.2
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox

from config import (
    APP_NAME,
    APP_SUBTITLE,
    APP_VERSION,
    COLOR_BG,
    COLOR_PANEL,
    COLOR_BORDER,
    COLOR_TEXT,
    COLOR_TEXT_MUTED,
    COLOR_GREEN,
    COLOR_GREEN_HOVER,
    COLOR_CYAN,
    FONT_TITLE,
    FONT_NORMAL,
    FONT_SUBTITLE,
)


class PantallaInicio(ctk.CTkFrame):
    """Pantalla principal de selección de modo de trabajo."""

    def __init__(self, master, app):
        super().__init__(master, fg_color=COLOR_BG)
        self.app = app
        self.pack(fill="both", expand=True)
        self._crear()

    def _crear(self):
        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PANEL, height=90, corner_radius=0)
        cabecera.pack(fill="x")

        ctk.CTkLabel(cabecera, text=APP_NAME, font=FONT_TITLE,
                     text_color=COLOR_GREEN).place(x=30, y=15)

        ctk.CTkLabel(cabecera, text=APP_SUBTITLE, font=FONT_NORMAL,
                     text_color=COLOR_TEXT_MUTED).place(x=32, y=58)

        ctk.CTkLabel(cabecera, text=APP_VERSION, font=FONT_NORMAL,
                     text_color=COLOR_CYAN).place(relx=0.97, y=35, anchor="e")

        panel = ctk.CTkFrame(
            self,
            width=520,
            height=360,
            fg_color=COLOR_PANEL,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        panel.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            panel,
            text="SELECCIONAR MODO DE TRABAJO",
            font=FONT_SUBTITLE,
            text_color=COLOR_TEXT,
        ).pack(pady=(40, 30))

        self._boton(panel, "CONTINUAR TEMPORADA", self.continuar).pack(pady=10)
        self._boton(panel, "NUEVA TEMPORADA", self.app.mostrar_nueva_temporada).pack(pady=10)
        self._boton(panel, "SALIR", self.app.cerrar_aplicacion).pack(pady=10)

    def _boton(self, master, texto, comando):
        return ctk.CTkButton(
            master,
            text=texto,
            command=comando,
            width=320,
            height=55,
            fg_color=COLOR_GREEN,
            hover_color=COLOR_GREEN_HOVER,
            text_color="black",
            font=("Segoe UI", 17, "bold"),
        )

    def continuar(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )

        if not archivo:
            return

        self.app.contexto.archivo_actual = archivo

        messagebox.showinfo(
            "Continuar temporada",
            f"Archivo seleccionado:\n\n{archivo}\n\nLa lectura del Excel llegará en la v0.3."
        )

        self.app.mostrar_medido()
