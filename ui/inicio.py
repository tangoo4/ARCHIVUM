"""
Pantalla de inicio de Archivum
Versión 0.5.2

Cambios:
- Al continuar temporada pide la medida estándar.
- Intenta deducir tipo, notario y año desde el nombre del archivo.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from pathlib import Path

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
    DIR_TEMPORADAS,
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
            text_color="black",
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

        medida = simpledialog.askstring(
            "Medida estándar",
            "Introduce la medida estándar de esta temporada:",
            initialvalue=self.app.contexto.medida_estandar or "8,5",
            parent=self,
        )

        if not medida:
            return

        medida = medida.strip().replace(".", ",")

        if not self._validar_medida(medida):
            messagebox.showwarning(
                "Medida incorrecta",
                "La medida estándar debe ser un número máximo 10. Usa coma si hay decimal."
            )
            return

        self.app.contexto.archivo_actual = archivo
        self.app.contexto.medida_estandar = medida

        self._rellenar_contexto_desde_nombre(archivo)

        self.app.mostrar_medido()

    def _validar_medida(self, medida):
        try:
            valor = float(medida.replace(",", "."))
        except ValueError:
            return False

        return 0 < valor <= 10

    def _rellenar_contexto_desde_nombre(self, archivo):
        """
        Intenta deducir tipo, notario y año desde nombres tipo:
        PROTOCOLO_PEREZ_2025.xlsx
        POLIZAS_PEREZ_2025.xlsx
        LIBRO_INDICADOR_PEREZ_2025.xlsx
        """

        nombre = Path(archivo).stem
        partes = nombre.split("_")

        if len(partes) < 3:
            return

        anio = partes[-1]

        if not anio.isdigit():
            return

        if partes[0] == "LIBRO" and len(partes) >= 4 and partes[1] == "INDICADOR":
            tipo = "LIBRO INDICADOR"
            notario = "_".join(partes[2:-1])
        else:
            tipo = partes[0]
            notario = "_".join(partes[1:-1])

        self.app.contexto.tipo = tipo
        self.app.contexto.notario = notario
        self.app.contexto.anio = anio
