"""
==========================================================
ARCHIVUM
Pantalla: Nueva Temporada
Versión: 0.3.0
==========================================================

Crea físicamente un Excel nuevo desde la plantilla de medido.

Flujo:
- Validar datos.
- Generar nombre del archivo.
- Copiar la plantilla.
- Guardar el Excel en la carpeta de trabajo actual.
- Actualizar el contexto.
- Abrir pantalla de medido.

IMPORTANTE:
La plantilla debe estar en:
ARCHIVUM/excel/PLANTILLA_MEDIDO.xlsx

Si no existe, intentará usar:
ARCHIVUM/excel/plantilla.xlsx
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import shutil
import unicodedata

from config import (
    COLOR_BG,
    COLOR_PANEL,
    COLOR_PANEL_2,
    COLOR_BORDER,
    COLOR_TEXT,
    COLOR_TEXT_MUTED,
    COLOR_GREEN,
    COLOR_GREEN_HOVER,
    COLOR_CYAN,
    FONT_TITLE,
    FONT_SUBTITLE,
    FONT_NORMAL,
    FONT_BUTTON,
    ENTRY_WIDTH,
    ENTRY_HEIGHT,
)


def limpiar_nombre(texto: str) -> str:
    """
    Convierte un texto en formato seguro para nombre de archivo.

    Ejemplo:
    Salvador Farrés Ripoll -> SALVADOR_FARRES_RIPOLL
    """

    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = texto.replace(" ", "_")
    texto = "".join(c for c in texto if c.isalnum() or c == "_")
    return texto


class PantallaNuevaTemporada(ctk.CTkFrame):
    """
    Pantalla para crear una nueva temporada desde plantilla.
    """

    def __init__(self, master, app):
        super().__init__(master, fg_color=COLOR_BG)

        self.app = app

        self.pack(fill="both", expand=True)
        self._crear_interfaz()

    # ======================================================
    # INTERFAZ
    # ======================================================

    def _crear_interfaz(self):
        self._crear_cabecera()
        self._crear_panel_formulario()

    def _crear_cabecera(self):
        cabecera = ctk.CTkFrame(
            self,
            fg_color=COLOR_PANEL,
            height=90,
            corner_radius=0,
        )
        cabecera.pack(fill="x")

        ctk.CTkLabel(
            cabecera,
            text="NUEVA TEMPORADA",
            font=FONT_TITLE,
            text_color=COLOR_GREEN,
        ).place(x=30, y=15)

        ctk.CTkLabel(
            cabecera,
            text="Crear un nuevo archivo Excel desde plantilla",
            font=FONT_NORMAL,
            text_color=COLOR_TEXT_MUTED,
        ).place(x=32, y=58)

        ctk.CTkLabel(
            cabecera,
            text="v0.3.0",
            font=FONT_NORMAL,
            text_color=COLOR_CYAN,
        ).place(relx=0.97, y=35, anchor="e")

    def _crear_panel_formulario(self):
        panel = ctk.CTkFrame(
            self,
            width=640,
            height=540,
            fg_color=COLOR_PANEL,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        panel.place(relx=0.5, rely=0.54, anchor="center")

        ctk.CTkLabel(
            panel,
            text="DATOS DE TEMPORADA",
            font=FONT_SUBTITLE,
            text_color=COLOR_TEXT,
        ).pack(pady=(35, 25))

        self.tipo = self._crear_combo(
            panel,
            "TIPO",
            ["PROTOCOLO", "POLIZAS", "LIBRO INDICADOR"],
        )

        self.notario = self._crear_entrada(panel, "NOTARIO")
        self.anio = self._crear_entrada(panel, "AÑO")
        self.medida = self._crear_entrada(panel, "MEDIDA ESTÁNDAR")

        self._configurar_enter()

        botones = ctk.CTkFrame(panel, fg_color="transparent")
        botones.pack(pady=(20, 0))

        self._boton_secundario(
            botones,
            "VOLVER",
            self.app.mostrar_inicio,
        ).pack(side="left", padx=12)

        self._boton_principal(
            botones,
            "CREAR TEMPORADA",
            self._crear_temporada,
        ).pack(side="left", padx=12)

        self.notario.focus_set()

    def _crear_combo(self, master, etiqueta, valores):
        ctk.CTkLabel(
            master,
            text=etiqueta,
            font=FONT_NORMAL,
            text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=140)

        combo = ctk.CTkComboBox(
            master,
            values=valores,
            width=ENTRY_WIDTH,
            height=ENTRY_HEIGHT,
            fg_color=COLOR_PANEL_2,
            border_color=COLOR_BORDER,
            button_color=COLOR_GREEN,
            button_hover_color=COLOR_GREEN_HOVER,
            text_color=COLOR_TEXT,
            dropdown_fg_color=COLOR_PANEL_2,
            dropdown_text_color=COLOR_TEXT,
            font=FONT_NORMAL,
            dropdown_font=FONT_NORMAL,
        )
        combo.set(valores[0])
        combo.pack(pady=(5, 18))

        return combo

    def _crear_entrada(self, master, etiqueta):
        ctk.CTkLabel(
            master,
            text=etiqueta,
            font=FONT_NORMAL,
            text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=140)

        campo = ctk.CTkEntry(
            master,
            width=ENTRY_WIDTH,
            height=ENTRY_HEIGHT,
            fg_color=COLOR_PANEL_2,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=FONT_NORMAL,
            corner_radius=8,
        )
        campo.pack(pady=(5, 18))

        campo.bind("<KeyRelease>", self._mayusculas)

        return campo

    def _boton_principal(self, master, texto, comando):
        return ctk.CTkButton(
            master,
            text=texto,
            command=comando,
            width=260,
            height=50,
            fg_color=COLOR_GREEN,
            hover_color=COLOR_GREEN_HOVER,
            text_color="black",
            font=FONT_BUTTON,
            corner_radius=12,
        )

    def _boton_secundario(self, master, texto, comando):
        return ctk.CTkButton(
            master,
            text=texto,
            command=comando,
            width=180,
            height=50,
            fg_color=COLOR_PANEL_2,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=FONT_NORMAL,
            corner_radius=12,
        )

    # ======================================================
    # COMPORTAMIENTO
    # ======================================================

    def _configurar_enter(self):
        self.notario.bind("<Return>", lambda _event: self.anio.focus_set())
        self.anio.bind("<Return>", lambda _event: self.medida.focus_set())
        self.medida.bind("<Return>", lambda _event: self._crear_temporada())

    def _mayusculas(self, event):
        campo = event.widget
        texto = campo.get()
        cursor = campo.index("insert")

        campo.delete(0, "end")
        campo.insert(0, texto.upper())
        campo.icursor(cursor)

    def _validar(self):
        notario = self.notario.get().strip()
        anio = self.anio.get().strip()
        medida = self.medida.get().strip()

        if not notario:
            return False, "Introduce el nombre del notario."

        if not anio.isdigit() or len(anio) != 4:
            return False, "El año debe tener 4 dígitos."

        if not self._validar_medida(medida):
            return False, "La medida debe ser un número máximo 10. Usa coma si hay decimal."

        return True, ""

    def _validar_medida(self, medida: str) -> bool:
        if not medida:
            return False

        try:
            valor = float(medida.replace(",", "."))
        except ValueError:
            return False

        return 0 < valor <= 10

    # ======================================================
    # CREACIÓN DE EXCEL
    # ======================================================

    def _buscar_plantilla(self) -> Path | None:
        """
        Busca la plantilla en la carpeta excel/.
        Acepta dos nombres para evitar errores:
        - PLANTILLA_MEDIDO.xlsx
        - plantilla.xlsx
        """

        base = Path.cwd()
        posibles = [
            base / "excel" / "PLANTILLA_MEDIDO.xlsx",
            base / "excel" / "plantilla.xlsx",
        ]

        for ruta in posibles:
            if ruta.exists():
                return ruta

        return None

    def _crear_temporada(self):
        ok, mensaje = self._validar()

        if not ok:
            messagebox.showwarning("Datos incorrectos", mensaje)
            return

        plantilla = self._buscar_plantilla()

        if plantilla is None:
            messagebox.showerror(
                "Plantilla no encontrada",
                (
                    "No se encontró la plantilla.\n\n"
                    "Debe estar en una de estas rutas:\n\n"
                    "ARCHIVUM/excel/PLANTILLA_MEDIDO.xlsx\n"
                    "ARCHIVUM/excel/plantilla.xlsx"
                ),
            )
            return

        tipo = self.tipo.get().strip().upper()
        tipo_archivo = limpiar_nombre(tipo)
        notario_limpio = limpiar_nombre(self.notario.get())
        anio = self.anio.get().strip()
        medida = self.medida.get().strip()

        nombre_archivo = f"{tipo_archivo}_{notario_limpio}_{anio}.xlsx"

        # Por ahora se crea en la carpeta raíz desde donde se ejecuta Archivum.
        # Esto encaja con la idea de ejecutar Archivum dentro de la carpeta del notario.
        destino = Path.cwd() / nombre_archivo

        if destino.exists():
            respuesta = messagebox.askyesno(
                "Archivo existente",
                (
                    f"Ya existe el archivo:\n\n{destino}\n\n"
                    "¿Deseas sobrescribirlo?"
                ),
            )

            if not respuesta:
                return

        try:
            shutil.copy2(plantilla, destino)
        except Exception as e:
            messagebox.showerror(
                "Error al crear archivo",
                f"No se pudo copiar la plantilla:\n\n{e}",
            )
            return

        self.app.contexto.tipo = tipo
        self.app.contexto.notario = notario_limpio
        self.app.contexto.anio = anio
        self.app.contexto.medida_estandar = medida
        self.app.contexto.archivo_actual = str(destino)
        self.app.contexto.tomo_actual = 1

        messagebox.showinfo(
            "Temporada creada",
            (
                "Temporada creada correctamente.\n\n"
                f"Archivo:\n{destino}\n\n"
                f"Medida estándar: {medida}"
            ),
        )

        self.app.mostrar_medido()
