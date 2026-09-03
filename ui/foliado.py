"""Formulario independiente para introducir el foliado tomo a tomo."""

from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    COLOR_BG,
    COLOR_BORDER,
    COLOR_CYAN,
    COLOR_GREEN,
    COLOR_GREEN_HOVER,
    COLOR_PANEL,
    COLOR_PANEL_2,
    COLOR_TEXT,
    COLOR_TEXT_MUTED,
    FONT_NORMAL,
    FONT_SUBTITLE,
)
from excel.gestor_foliado import guardar_foliado, leer_foliado


class VentanaFoliado(ctk.CTkToplevel):
    """Permite completar o corregir el foliado sin entrar en el medido."""

    def __init__(self, master, ruta_excel, al_guardar=None):
        super().__init__(master)
        self.ruta_excel = ruta_excel
        self.al_guardar = al_guardar
        self.tomos = []
        self.indice_actual = 0

        self.title("Añadir foliado")
        self.geometry("720x560")
        self.minsize(650, 500)
        self.configure(fg_color=COLOR_BG)
        self.transient(master.winfo_toplevel())
        self.grab_set()

        self._crear_interfaz()
        self._cargar_tomos()

    def _crear_interfaz(self):
        ctk.CTkLabel(
            self,
            text="AÑADIR FOLIADO",
            font=FONT_SUBTITLE,
            text_color=COLOR_GREEN,
        ).pack(anchor="w", padx=25, pady=(22, 4))

        ctk.CTkLabel(
            self,
            text="Introduce el número de la última hoja de cada tomo.",
            font=FONT_NORMAL,
            text_color=COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=25, pady=(0, 16))

        panel_tabla = ctk.CTkFrame(self, fg_color=COLOR_PANEL, corner_radius=12)
        panel_tabla.pack(fill="both", expand=True, padx=25, pady=(0, 15))

        columnas = ("tomo", "foliado", "hojas", "estado")
        self.tabla = ttk.Treeview(panel_tabla, columns=columnas, show="headings", height=10)
        for columna, titulo, ancho in (
            ("tomo", "TOMO", 90),
            ("foliado", "Nº HOJA FINAL", 160),
            ("hojas", "Nº HOJAS TOMO", 170),
            ("estado", "ESTADO", 130),
        ):
            self.tabla.heading(columna, text=titulo)
            self.tabla.column(columna, width=ancho, anchor="center")
        self.tabla.pack(fill="both", expand=True, padx=15, pady=15)
        self.tabla.bind("<<TreeviewSelect>>", self._seleccionar_desde_tabla)

        panel = ctk.CTkFrame(self, fg_color=COLOR_PANEL, corner_radius=12, height=125)
        panel.pack(fill="x", padx=25, pady=(0, 22))
        panel.pack_propagate(False)

        self.label_tomo = ctk.CTkLabel(panel, text="TOMO -", font=FONT_SUBTITLE, text_color=COLOR_CYAN)
        self.label_tomo.place(x=22, y=18)
        self.label_anterior = ctk.CTkLabel(panel, text="Foliado anterior: —", font=FONT_NORMAL, text_color=COLOR_TEXT_MUTED)
        self.label_anterior.place(x=22, y=55)
        self.label_resultado = ctk.CTkLabel(panel, text="Hojas del tomo: —", font=FONT_NORMAL, text_color=COLOR_TEXT)
        self.label_resultado.place(x=22, y=84)

        self.entrada = ctk.CTkEntry(
            panel, width=170, height=40, fg_color=COLOR_PANEL_2,
            border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=FONT_NORMAL,
            placeholder_text="Nº hoja final",
        )
        self.entrada.place(relx=0.56, y=25, anchor="n")
        self.entrada.bind("<KeyRelease>", self._actualizar_resultado)
        self.entrada.bind("<Return>", lambda _event: self._guardar())

        ctk.CTkButton(
            panel, text="GUARDAR Y SIGUIENTE", command=self._guardar,
            width=190, height=40, fg_color=COLOR_GREEN,
            hover_color=COLOR_GREEN_HOVER, text_color="black",
            font=("Segoe UI", 13, "bold"),
        ).place(relx=0.82, y=25, anchor="n")
        ctk.CTkButton(
            panel, text="CERRAR", command=self.destroy,
            width=120, height=32, fg_color=COLOR_PANEL_2,
            hover_color=COLOR_BORDER, text_color=COLOR_TEXT, font=FONT_NORMAL,
        ).place(relx=0.82, y=76, anchor="n")

    def _cargar_tomos(self, seleccionar=None):
        try:
            self.tomos = leer_foliado(self.ruta_excel)
        except Exception as exc:
            messagebox.showerror("Error al abrir la temporada", str(exc), parent=self)
            self.destroy()
            return

        for item in self.tabla.get_children():
            self.tabla.delete(item)

        for i, tomo in enumerate(self.tomos):
            completo = tomo["foliado"] not in (None, "")
            self.tabla.insert(
                "", "end", iid=str(i),
                values=(
                    tomo["tomo"],
                    tomo["foliado"] if completo else "",
                    tomo["hojas"] if completo else "",
                    "COMPLETO" if completo else "PENDIENTE",
                ),
            )

        if not self.tomos:
            messagebox.showinfo("Sin tomos", "La temporada todavía no contiene tomos.", parent=self)
            self.destroy()
            return

        if seleccionar is None:
            seleccionar = next((i for i, tomo in enumerate(self.tomos) if tomo["foliado"] in (None, "")), len(self.tomos) - 1)
        self._seleccionar_indice(seleccionar)

    def _seleccionar_desde_tabla(self, _event=None):
        seleccion = self.tabla.selection()
        if seleccion:
            self._seleccionar_indice(int(seleccion[0]), actualizar_tabla=False)

    def _seleccionar_indice(self, indice, actualizar_tabla=True):
        self.indice_actual = max(0, min(indice, len(self.tomos) - 1))
        tomo = self.tomos[self.indice_actual]
        anterior = self.tomos[self.indice_actual - 1]["foliado"] if self.indice_actual > 0 else 0

        self.label_tomo.configure(text=f"TOMO {tomo['tomo']}")
        self.label_anterior.configure(text=f"Foliado anterior: {anterior if anterior not in (None, '') else 'PENDIENTE'}")
        self.entrada.delete(0, "end")
        if tomo["foliado"] not in (None, ""):
            self.entrada.insert(0, str(tomo["foliado"]))
        self._actualizar_resultado()
        self.entrada.focus_set()

        if actualizar_tabla:
            self.tabla.selection_set(str(self.indice_actual))
            self.tabla.focus(str(self.indice_actual))
            self.tabla.see(str(self.indice_actual))

    def _actualizar_resultado(self, _event=None):
        texto = self.entrada.get().strip()
        try:
            final = int(texto)
            anterior = int(self.tomos[self.indice_actual - 1]["foliado"]) if self.indice_actual > 0 else 0
            hojas = final - anterior
            resultado = str(hojas) if hojas > 0 else "NO VÁLIDO"
        except (TypeError, ValueError):
            resultado = "—"
        self.label_resultado.configure(text=f"Hojas del tomo: {resultado}")

    def _guardar(self):
        valor = self.entrada.get().strip()
        if not valor.isdigit():
            messagebox.showwarning("Foliado incorrecto", "Introduce un número entero mayor que cero.", parent=self)
            self.entrada.focus_set()
            return

        tomo = self.tomos[self.indice_actual]["tomo"]
        try:
            guardar_foliado(self.ruta_excel, tomo, int(valor))
        except Exception as exc:
            messagebox.showwarning("No se puede guardar", str(exc), parent=self)
            self.entrada.focus_set()
            return

        siguiente = min(self.indice_actual + 1, len(self.tomos) - 1)
        self._cargar_tomos(seleccionar=siguiente)
        if self.al_guardar:
            self.al_guardar()
