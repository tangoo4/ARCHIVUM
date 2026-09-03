"""
==========================================================
ARCHIVUM
Configuración general de la aplicación
Versión: 0.7.0
==========================================================

Modo local portable:

ARCHIVUM/
├── ARCHIVUM.exe
└── temporadas/

La carpeta temporadas se crea siempre junto al ejecutable.
"""

from pathlib import Path
import sys


# ==========================================================
# INFORMACIÓN DE LA APLICACIÓN
# ==========================================================

APP_NAME = "ARCHIVUM"
APP_SUBTITLE = "Sistema de Gestión de Encuadernación Notarial"
APP_VERSION = "0.7.0"
APP_BUILD = 1
APP_AUTHOR = "Iván"


# ==========================================================
# UBICACIÓN DE LA APLICACIÓN
# ==========================================================


def get_app_dir() -> Path:
    """
    Carpeta donde está el EXE o el proyecto en desarrollo.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def get_resource_dir() -> Path:
    """
    Carpeta de recursos.

    En EXE de PyInstaller, los recursos internos van en sys._MEIPASS.
    En desarrollo, van en la carpeta del proyecto.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)

    return Path(__file__).resolve().parent


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
# RUTAS
# ==========================================================

BASE_DIR = get_app_dir()
RESOURCE_DIR = get_resource_dir()
WORK_ROOT = BASE_DIR

# Carpetas de código en desarrollo
DIR_CORE = BASE_DIR / "core"
DIR_UI = BASE_DIR / "ui"

# Carpeta de trabajo, siempre junto a la aplicación
DIR_TEMPORADAS = BASE_DIR / "temporadas"

# Rutas reservadas para funciones futuras, también locales
DIR_EXCEL = BASE_DIR / "excel"
DIR_ASSETS = BASE_DIR / "assets"
DIR_LOGS = BASE_DIR / "logs"
DIR_BACKUPS = BASE_DIR / "backups"
DIR_DATOS = BASE_DIR / "datos"
DIR_TEMP = BASE_DIR / "temp"

# Recursos externos opcionales en desarrollo
LOGO_PATH = DIR_ASSETS / "logo.png"
ICON_PATH = DIR_ASSETS / "archivum.ico"

# Recursos internos del EXE / proyecto
BUNDLED_EXCEL_DIR = RESOURCE_DIR / "excel"
BUNDLED_PLANTILLA_EXCEL = BUNDLED_EXCEL_DIR / "PLANTILLA_MEDIDO.xlsx"
BUNDLED_PLANTILLA_EXCEL_ALT = BUNDLED_EXCEL_DIR / "plantilla.xlsx"

# Plantilla externa opcional junto al proyecto en desarrollo
PLANTILLA_EXCEL = DIR_EXCEL / "PLANTILLA_MEDIDO.xlsx"
PLANTILLA_EXCEL_ALT = DIR_EXCEL / "plantilla.xlsx"

ULTIMO_ARCHIVO_PATH = DIR_DATOS / "ultimo_archivo.txt"
HISTORIAL_PATH = DIR_DATOS / "historial.txt"


# ==========================================================
# COLORES
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
# COMPONENTES
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
# BACKUPS / FUTURO
# ==========================================================

BACKUP_ON_CLOSE = False
BACKUP_FOLDER_NAME = "BACKUPS"
BACKUP_DATETIME_FORMAT = "%Y%m%d_%H%M%S"


# ==========================================================
# UTILIDADES
# ==========================================================

def ensure_project_dirs() -> None:
    """
    Crea la carpeta de temporadas junto al EXE o al proyecto.
    """
    DIR_TEMPORADAS.mkdir(parents=True, exist_ok=True)


def get_plantilla_path() -> Path | None:
    """
    Devuelve la plantilla disponible.

    Prioridad:
    1. Plantilla incluida dentro del EXE.
    2. Plantilla externa en la carpeta excel del proyecto.

    Así, en la versión final, el usuario no necesita tener ni tocar
    la carpeta excel ni la plantilla.
    """
    if BUNDLED_PLANTILLA_EXCEL.exists():
        return BUNDLED_PLANTILLA_EXCEL

    if BUNDLED_PLANTILLA_EXCEL_ALT.exists():
        return BUNDLED_PLANTILLA_EXCEL_ALT

    if PLANTILLA_EXCEL.exists():
        return PLANTILLA_EXCEL

    if PLANTILLA_EXCEL_ALT.exists():
        return PLANTILLA_EXCEL_ALT

    return None


def modo_trabajo_texto() -> str:
    """
    Texto breve para saber dónde está trabajando Archivum.
    """
    return f"Carpeta de la aplicación: {WORK_ROOT}"
