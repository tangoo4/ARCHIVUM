"""
==========================================================
ARCHIVUM
Clase principal de la aplicación
Versión: 0.2.0
==========================================================

Este módulo contiene la clase ArchivumApp, que actúa como
controlador principal de toda la aplicación.

Responsabilidades:
- Crear la ventana principal.
- Aplicar configuración general.
- Mantener el contexto de trabajo.
- Cambiar entre pantallas.
- Gestionar el cierre del programa.
"""

import customtkinter as ctk
from dataclasses import dataclass

from config import (
    APP_NAME,
    APP_VERSION,
    WINDOW_WIDTH,
    WINDOW_HEIGHT,
    WINDOW_MIN_WIDTH,
    WINDOW_MIN_HEIGHT,
    START_MAXIMIZED,
    RESIZABLE,
    COLOR_BG,
    ensure_project_dirs,
)


# ==========================================================
# CONTEXTO DE TRABAJO
# ==========================================================

@dataclass
class AppContext:
    """
    Guarda el estado actual de la aplicación.

    De momento solo contiene datos básicos, pero más adelante
    se usará para compartir información entre pantallas:
    archivo abierto, notario, año, medida estándar, tomo actual, etc.
    """

    archivo_actual: str | None = None
    tipo: str = ""
    notario: str = ""
    anio: str = ""
    medida_estandar: str = "8,5"
    tomo_actual: int = 1


# ==========================================================
# APLICACIÓN PRINCIPAL
# ==========================================================

class ArchivumApp(ctk.CTk):
    """
    Ventana principal de Archivum.

    Esta clase no debe contener lógica de Excel ni lógica específica
    de una pantalla concreta. Solo debe coordinar la aplicación.
    """

    def __init__(self):
        super().__init__()

        ensure_project_dirs()

        self.contexto = AppContext()
        self.pantalla_actual = None

        self.configurar_ventana()
        self.mostrar_inicio()

    # ======================================================
    # CONFIGURACIÓN DE VENTANA
    # ======================================================

    def configurar_ventana(self):
        """
        Aplica la configuración principal de la ventana.
        """

        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.resizable(RESIZABLE, RESIZABLE)
        self.configure(fg_color=COLOR_BG)

        if START_MAXIMIZED:
            try:
                self.state("zoomed")
            except Exception:
                pass

        self.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)

    # ======================================================
    # GESTIÓN DE PANTALLAS
    # ======================================================

    def limpiar_pantalla(self):
        """
        Destruye la pantalla actual antes de cargar una nueva.
        """

        if self.pantalla_actual is not None:
            self.pantalla_actual.destroy()
            self.pantalla_actual = None

    def mostrar_inicio(self):
        """
        Muestra la pantalla de inicio.

        En v0.2 importamos las pantallas dentro del método para evitar
        dependencias circulares mientras reorganizamos el proyecto.
        """

        from ui.inicio import PantallaInicio

        self.limpiar_pantalla()

        self.pantalla_actual = PantallaInicio(
            master=self,
            app=self
        )

    def mostrar_nueva_temporada(self):
        """
        Muestra la pantalla de creación de nueva temporada.
        """

        from ui.nueva_temporada import PantallaNuevaTemporada

        self.limpiar_pantalla()

        self.pantalla_actual = PantallaNuevaTemporada(
            master=self,
            app=self
        )

    def mostrar_medido(self):
        """
        Muestra la pantalla principal de medido.
        """

        from ui.medido import PantallaMedido

        self.limpiar_pantalla()

        self.pantalla_actual = PantallaMedido(
            master=self,
            app=self
        )

    # ======================================================
    # CIERRE
    # ======================================================

    def cerrar_aplicacion(self):
        """
        Cierra Archivum.

        Más adelante aquí se añadirá:
        - confirmación si hay cambios pendientes,
        - backup automático,
        - registro en log.
        """

        self.destroy()
