"""Formulario para corregir los datos de un tomo ya guardado."""

from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    COLOR_BG, COLOR_BORDER, COLOR_CYAN, COLOR_GREEN, COLOR_GREEN_HOVER,
    COLOR_PANEL, COLOR_PANEL_2, COLOR_TEXT, COLOR_TEXT_MUTED,
    FONT_NORMAL, FONT_SUBTITLE, MAX_MEDIDA, MAX_OBSERVACIONES,
)
from excel.gestor_tomos import comprobar_continuidad, leer_tomos, modificar_tomo


class VentanaModificarTomo(ctk.CTkToplevel):
    """Selecciona un tomo, permite editar A:H y conserva J:K."""

    CAMPOS = (
        ("anio", "AÑO"),
        ("matriz_inicio", "MATRIZ INICIO"),
        ("fecha_inicio", "FECHA INICIO"),
        ("matriz_final", "MATRIZ FINAL"),
        ("fecha_final", "FECHA FINAL"),
        ("medida", "MEDIDA"),
        ("observaciones", "OBSERVACIONES"),
    )

    def __init__(self, master, ruta_excel, medida_estandar, al_guardar=None):
        super().__init__(master)
        self.ruta_excel = ruta_excel
        self.medida_estandar = medida_estandar
        self.al_guardar = al_guardar
        self.tomos = []
        self.indice_actual = 0
        self.entradas = {}

        self.title("Modificar tomo")
        self.geometry("980x650")
        self.minsize(900, 600)
        self.configure(fg_color=COLOR_BG)
        self.transient(master.winfo_toplevel())
        self.grab_set()

        self._crear_interfaz()
        self._cargar_tomos()

    def _crear_interfaz(self):
        ctk.CTkLabel(self, text="MODIFICAR TOMO", font=FONT_SUBTITLE, text_color=COLOR_GREEN).pack(
            anchor="w", padx=25, pady=(22, 4)
        )
        ctk.CTkLabel(
            self, text="Selecciona un tomo y corrige únicamente los datos necesarios.",
            font=FONT_NORMAL, text_color=COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=25, pady=(0, 15))

        cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=25, pady=(0, 15))

        panel_lista = ctk.CTkFrame(cuerpo, width=360, fg_color=COLOR_PANEL, corner_radius=12)
        panel_lista.pack(side="left", fill="y", padx=(0, 15))
        panel_lista.pack_propagate(False)

        self.tabla = ttk.Treeview(panel_lista, columns=("tomo", "rango"), show="headings")
        self.tabla.heading("tomo", text="TOMO")
        self.tabla.heading("rango", text="MATRICES")
        self.tabla.column("tomo", width=80, anchor="center")
        self.tabla.column("rango", width=230, anchor="center")
        self.tabla.pack(fill="both", expand=True, padx=15, pady=15)
        self.tabla.bind("<<TreeviewSelect>>", self._seleccionar)

        panel_form = ctk.CTkFrame(cuerpo, fg_color=COLOR_PANEL, corner_radius=12)
        panel_form.pack(side="left", fill="both", expand=True)

        self.label_tomo = ctk.CTkLabel(panel_form, text="TOMO -", font=FONT_SUBTITLE, text_color=COLOR_CYAN)
        self.label_tomo.pack(anchor="w", padx=28, pady=(22, 15))

        for clave, etiqueta in self.CAMPOS:
            fila = ctk.CTkFrame(panel_form, fg_color="transparent")
            fila.pack(fill="x", padx=28, pady=5)
            ctk.CTkLabel(fila, text=etiqueta, width=165, anchor="w", font=FONT_NORMAL, text_color=COLOR_TEXT).pack(side="left")
            entrada = ctk.CTkEntry(
                fila, height=36, fg_color=COLOR_PANEL_2, border_color=COLOR_BORDER,
                text_color=COLOR_TEXT, font=FONT_NORMAL,
            )
            entrada.pack(side="left", fill="x", expand=True)
            self.entradas[clave] = entrada

        botones = ctk.CTkFrame(panel_form, fg_color="transparent")
        botones.pack(fill="x", padx=28, pady=(20, 15))
        ctk.CTkButton(
            botones, text="CANCELAR", command=self.destroy, width=140, height=42,
            fg_color=COLOR_PANEL_2, hover_color=COLOR_BORDER, text_color=COLOR_TEXT, font=FONT_NORMAL,
        ).pack(side="right", padx=(10, 0))
        ctk.CTkButton(
            botones, text="GUARDAR CAMBIOS", command=self._guardar, width=190, height=42,
            fg_color=COLOR_GREEN, hover_color=COLOR_GREEN_HOVER, text_color="black",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="right")

    def _cargar_tomos(self, seleccionar=0):
        try:
            self.tomos = leer_tomos(self.ruta_excel)
        except Exception as exc:
            messagebox.showerror("Error al abrir la temporada", str(exc), parent=self)
            self.destroy()
            return

        if not self.tomos:
            messagebox.showinfo("Sin tomos", "La temporada todavía no contiene tomos.", parent=self)
            self.destroy()
            return

        for item in self.tabla.get_children():
            self.tabla.delete(item)
        for i, tomo in enumerate(self.tomos):
            rango = f"{tomo['matriz_inicio']} - {tomo['matriz_final']}"
            self.tabla.insert("", "end", iid=str(i), values=(tomo["tomo"], rango))

        self.tabla.selection_set(str(seleccionar))
        self.tabla.focus(str(seleccionar))
        self.tabla.see(str(seleccionar))
        self._mostrar_tomo(seleccionar)

    def _seleccionar(self, _event=None):
        seleccion = self.tabla.selection()
        if seleccion:
            self._mostrar_tomo(int(seleccion[0]))

    def _mostrar_tomo(self, indice):
        self.indice_actual = indice
        tomo = self.tomos[indice]
        self.label_tomo.configure(text=f"TOMO {tomo['tomo']}")
        for clave, _ in self.CAMPOS:
            entrada = self.entradas[clave]
            entrada.delete(0, "end")
            valor = tomo[clave]
            if valor not in (None, ""):
                entrada.insert(0, str(valor))

    def _validar(self):
        datos = {clave: self.entradas[clave].get().strip().upper() for clave, _ in self.CAMPOS}
        obligatorios = ("anio", "matriz_inicio", "fecha_inicio", "matriz_final", "fecha_final", "medida")
        if any(not datos[clave] for clave in obligatorios):
            return None, "Completa todos los campos excepto observaciones."
        if not datos["matriz_inicio"].isdigit() or not datos["matriz_final"].isdigit():
            return None, "Las matrices inicial y final deben ser numéricas."
        if int(datos["matriz_final"]) < int(datos["matriz_inicio"]):
            return None, "La matriz final no puede ser menor que la matriz inicial."
        if len(datos["observaciones"]) > MAX_OBSERVACIONES:
            return None, f"Las observaciones no pueden superar {MAX_OBSERVACIONES} caracteres."
        if datos["medida"] != "?":
            try:
                medida = float(datos["medida"].replace(",", "."))
            except ValueError:
                return None, "La medida debe ser un número o ?."
            if medida <= 0 or medida > MAX_MEDIDA:
                return None, f"La medida debe estar entre 0 y {MAX_MEDIDA}."
            datos["medida"] = datos["medida"].replace(".", ",")
        return datos, ""

    def _guardar(self):
        datos, error = self._validar()
        if error:
            messagebox.showwarning("Datos incorrectos", error, parent=self)
            return

        avisos = comprobar_continuidad(
            self.tomos, self.indice_actual,
            int(datos["matriz_inicio"]), int(datos["matriz_final"]),
        )
        if avisos:
            continuar = messagebox.askyesno(
                "La continuidad cambiará",
                "\n\n".join(avisos) + "\n\n¿Deseas guardar igualmente la corrección?",
                parent=self,
            )
            if not continuar:
                return

        tomo = self.tomos[self.indice_actual]["tomo"]
        try:
            modificar_tomo(self.ruta_excel, tomo, datos, self.medida_estandar)
        except Exception as exc:
            messagebox.showerror("No se pudo modificar", str(exc), parent=self)
            return

        self._cargar_tomos(seleccionar=self.indice_actual)
        if self.al_guardar:
            self.al_guardar()
        messagebox.showinfo("Tomo modificado", f"El tomo {tomo} se ha actualizado correctamente.", parent=self)

