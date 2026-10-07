"""
Pantalla de inicio de ARCHIVUM.

Continuar temporada:
- Intenta deducir tipo, notario y año desde el nombre del archivo.
- Pregunta Notario, Año y Medida estándar antes de abrir Medido.
- Los valores detectados pueden corregirse manualmente.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog

from core.nomenclatura import extraer_datos_desde_nombre
from core.validaciones import normalizar_anio, normalizar_medida

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
    COLOR_GREEN_TEXT,
    COLOR_GREEN_HOVER,
    COLOR_CYAN,
    FONT_TITLE,
    FONT_NORMAL,
    FONT_SUBTITLE,
    DIR_TEMPORADAS,
    MAX_MEDIDA,
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

        ctk.CTkLabel(
            cabecera,
            text=APP_NAME,
            font=FONT_TITLE,
            text_color=COLOR_GREEN
        ).place(x=30, y=15)

        ctk.CTkLabel(
            cabecera,
            text=APP_SUBTITLE,
            font=FONT_NORMAL,
            text_color=COLOR_TEXT_MUTED
        ).place(x=32, y=58)

        ctk.CTkLabel(
            cabecera,
            text=APP_VERSION,
            font=FONT_NORMAL,
            text_color=COLOR_CYAN
        ).place(relx=0.97, y=35, anchor="e")

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
            text_color=COLOR_GREEN_TEXT,
            font=("Segoe UI", 17, "bold"),
        )

    def continuar(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar temporada",
            initialdir=DIR_TEMPORADAS if DIR_TEMPORADAS.exists() else None,
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )

        if not archivo:
            return

        # Intentar deducir datos del nombre, pero nunca depender de que
        # la detección sea perfecta: el usuario los confirma/corrige.
        datos_nombre = extraer_datos_desde_nombre(archivo)

        if datos_nombre is not None:
            tipo_detectado, notario_detectado, anio_detectado = datos_nombre
        else:
            tipo_detectado = self.app.contexto.tipo or ""
            notario_detectado = self.app.contexto.notario or ""
            anio_detectado = self.app.contexto.anio or ""

        notario = simpledialog.askstring(
            "Notario",
            "Confirma o corrige el nombre del notario:",
            initialvalue=notario_detectado,
            parent=self,
        )
        if notario:
            notario = notario.strip()
        if not notario:
            messagebox.showwarning("Notario incorrecto", "Debes indicar el nombre del notario.")
            return

        anio = simpledialog.askstring(
            "Año",
            "Confirma o corrige el año de la temporada:",
            initialvalue=str(anio_detectado or ""),
            parent=self,
        )
        if anio is None:
            return

        try:
            anio = normalizar_anio(anio)
        except ValueError as exc:
            messagebox.showwarning("Año incorrecto", str(exc))
            return

        medida = simpledialog.askstring(
            "Medida estándar",
            "Introduce la medida estándar de esta temporada:",
            initialvalue=self.app.contexto.medida_estandar or "8,5",
            parent=self,
        )
        if medida is None:
            return

        try:
            medida = normalizar_medida(medida, MAX_MEDIDA)
        except ValueError as exc:
            messagebox.showwarning("Medida incorrecta", str(exc))
            return

        # El contexto queda completo ANTES de abrir Medido.
        self.app.contexto.archivo_actual = archivo
        self.app.contexto.tipo = tipo_detectado or self.app.contexto.tipo or ""
        self.app.contexto.notario = notario
        self.app.contexto.anio = anio
        self.app.contexto.medida_estandar = medida

        self.app.mostrar_medido()
