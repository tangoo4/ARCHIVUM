"""
==========================================================
ARCHIVUM
Configuración general de la aplicación
Versión: 0.7.0
==========================================================

Modo portable definitivo:

T:\
├── ARCHIVUM.exe
├── temporadas\
└── archivum_system\   (oculta + sistema)
    ├── assets\
    ├── excel\
    ├── datos\
    ├── logs\
    ├── backups\
    └── temp\

El usuario solo debería ver:
- ARCHIVUM.exe
- temporadas
"""

from pathlib import Path
import subprocess
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
# MODO PORTABLE USB
# ==========================================================

USB_DRIVE = Path("T:/")
SYSTEM_FOLDER_NAME = "archivum_system"


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


def get_work_root() -> Path:
    """
    Si existe T:, Archivum trabaja siempre en T:.

    Si no existe T:, trabaja junto al EXE o junto al proyecto.
    Esto permite probar en VS Code sin USB.
    """
    if USB_DRIVE.exists():
        return USB_DRIVE

    return get_app_dir()


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
WORK_ROOT = get_work_root()

# Carpeta técnica oculta
DIR_SYSTEM = WORK_ROOT / SYSTEM_FOLDER_NAME

# Carpetas de código en desarrollo
DIR_CORE = BASE_DIR / "core"
DIR_UI = BASE_DIR / "ui"

# Carpeta visible de trabajo
DIR_TEMPORADAS = WORK_ROOT / "temporadas"

# Carpetas técnicas internas
DIR_EXCEL = DIR_SYSTEM / "excel"
DIR_ASSETS = DIR_SYSTEM / "assets"
DIR_LOGS = DIR_SYSTEM / "logs"
DIR_BACKUPS = DIR_SYSTEM / "backups"
DIR_DATOS = DIR_SYSTEM / "datos"
DIR_TEMP = DIR_SYSTEM / "temp"

# Recursos externos opcionales dentro de archivum_system
LOGO_PATH = DIR_ASSETS / "logo.png"
ICON_PATH = DIR_ASSETS / "archivum.ico"

# Recursos internos del EXE / proyecto
BUNDLED_EXCEL_DIR = RESOURCE_DIR / "excel"
BUNDLED_PLANTILLA_EXCEL = BUNDLED_EXCEL_DIR / "PLANTILLA_MEDIDO.xlsx"
BUNDLED_PLANTILLA_EXCEL_ALT = BUNDLED_EXCEL_DIR / "plantilla.xlsx"

# Plantilla externa opcional dentro de archivum_system/excel
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

def _windows_hide_system_folder(path: Path) -> None:
    """
    Marca una carpeta como oculta y de sistema en Windows.

    No requiere permisos de administrador.
    Si falla, Archivum sigue funcionando igualmente.
    """
    if not sys.platform.startswith("win"):
        return

    try:
        subprocess.run(
            ["attrib", "+h", "+s", str(path)],
            check=False,
            shell=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass


def ensure_project_dirs() -> None:
    """
    Crea la estructura portable definitiva.

    Visible:
    - temporadas

    Oculta:
    - archivum_system
    """
    DIR_TEMPORADAS.mkdir(exist_ok=True)

    DIR_SYSTEM.mkdir(exist_ok=True)

    for folder in (
        DIR_EXCEL,
        DIR_ASSETS,
        DIR_LOGS,
        DIR_BACKUPS,
        DIR_DATOS,
        DIR_TEMP,
    ):
        folder.mkdir(parents=True, exist_ok=True)

    _windows_hide_system_folder(DIR_SYSTEM)


def get_plantilla_path() -> Path | None:
    """
    Devuelve la plantilla disponible.

    Prioridad:
    1. Plantilla incluida dentro del EXE.
    2. Plantilla externa en archivum_system/excel.

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
    if USB_DRIVE.exists():
        return f"USB detectado: {USB_DRIVE}"

    return f"Modo local: {WORK_ROOT}"
