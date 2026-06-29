"""
==========================================================
ARCHIVUM
Configuración general de la aplicación
Versión: 0.2.0
==========================================================
"""

from pathlib import Path


# ==========================================================
# INFORMACIÓN DE LA APLICACIÓN
# ==========================================================

APP_NAME = "ARCHIVUM"
APP_SUBTITLE = "Sistema de Gestión de Encuadernación Notarial"
APP_VERSION = "0.2.0"
APP_BUILD = 1
APP_AUTHOR = "Iván"


# ==========================================================
# VENTANA
# ==========================================================

WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 750
WINDOW_MIN_WIDTH = 1000
WINDOW_MIN_HEIGHT = 650

START_MAXIMIZED = True
RESIZABLE = True


# ==========================================================
# RUTAS DEL PROYECTO
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

DIR_CORE = BASE_DIR / "core"
DIR_UI = BASE_DIR / "ui"
DIR_EXCEL = BASE_DIR / "excel"
DIR_ASSETS = BASE_DIR / "assets"
DIR_LOGS = BASE_DIR / "logs"
DIR_BACKUPS = BASE_DIR / "backups"
DIR_DATOS = BASE_DIR / "datos"
DIR_TEMP = BASE_DIR / "temp"

LOGO_PATH = DIR_ASSETS / "logo.png"
ICON_PATH = DIR_ASSETS / "archivum.ico"

PLANTILLA_EXCEL = DIR_EXCEL / "plantilla.xlsx"

ULTIMO_ARCHIVO_PATH = DIR_DATOS / "ultimo_archivo.txt"
HISTORIAL_PATH = DIR_DATOS / "historial.txt"


# ==========================================================
# COLORES - TEMA OSCURO
# ==========================================================

COLOR_BG = "#121212"
COLOR_PANEL = "#1E1E1E"
COLOR_PANEL_2 = "#2A2A2A"
COLOR_BORDER = "#3A3A3A"

COLOR_TEXT = "#FFFFFF"
COLOR_TEXT_MUTED = "#B8B8B8"

COLOR_GREEN = "#00FF66"
COLOR_GREEN_HOVER = "#00CC55"

COLOR_CYAN = "#00E5FF"
COLOR_RED = "#FF3030"
COLOR_YELLOW = "#FFE600"

COLOR_ENTRY_BG = COLOR_PANEL_2
COLOR_ENTRY_BORDER = "#444444"
COLOR_ENTRY_ACTIVE = COLOR_CYAN

COLOR_TABLE_BG = "#181818"
COLOR_TABLE_SELECTED = "#003C5A"


# ==========================================================
# FUENTES
# ==========================================================

FONT_FAMILY = "Segoe UI"

FONT_TITLE = (FONT_FAMILY, 32, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 20, "bold")
FONT_NORMAL = (FONT_FAMILY, 15)
FONT_SMALL = (FONT_FAMILY, 13)
FONT_BUTTON = (FONT_FAMILY, 18, "bold")
FONT_TOMO = (FONT_FAMILY, 42, "bold")
FONT_MONO = ("Consolas", 16)


# ==========================================================
# TAMAÑOS DE COMPONENTES
# ==========================================================

BUTTON_MAIN_WIDTH = 340
BUTTON_MAIN_HEIGHT = 60

BUTTON_SECONDARY_WIDTH = 220
BUTTON_SECONDARY_HEIGHT = 45

ENTRY_WIDTH = 360
ENTRY_HEIGHT = 42

RADIUS_SMALL = 8
RADIUS_MEDIUM = 12
RADIUS_LARGE = 16


# ==========================================================
# MEDIDO
# ==========================================================

DEFAULT_MEDIDA = "8,5"
MAX_MEDIDA = 10
MAX_OBSERVACIONES = 100

# Colores según medida:
# Blanco = medida estándar
# Rojo = medida distinta
# Amarillo = "?"
MEDIDA_COLOR_STANDARD = "WHITE"
MEDIDA_COLOR_SPECIAL = "RED"
MEDIDA_COLOR_UNKNOWN = "YELLOW"


# ==========================================================
# EXCEL
# ==========================================================

EXCEL_EXTENSION = ".xlsx"

EXCEL_HEADER_ROW = 1
EXCEL_FIRST_DATA_ROW = 2

COL_TOMO = "A"
COL_ANY = "B"
COL_PROT_INICIAL = "C"
COL_DATA_INICIAL = "D"
COL_PROT_FINAL = "E"
COL_DATA_FINAL = "F"
COL_GRUIX = "G"
COL_OBSERVACIONS = "H"


# ==========================================================
# BACKUPS
# ==========================================================

BACKUP_ON_CLOSE = True
BACKUP_FOLDER_NAME = "BACKUPS"
BACKUP_DATETIME_FORMAT = "%Y%m%d_%H%M%S"


# ==========================================================
# UTILIDADES
# ==========================================================

def ensure_project_dirs() -> None:
    """
    Crea las carpetas internas necesarias si no existen.
    """
    for folder in (
        DIR_CORE,
        DIR_UI,
        DIR_EXCEL,
        DIR_ASSETS,
        DIR_LOGS,
        DIR_BACKUPS,
        DIR_DATOS,
        DIR_TEMP,
    ):
        folder.mkdir(exist_ok=True)
