"""Reglas de nomenclatura para temporadas y archivos de ARCHIVUM."""

from pathlib import Path


TIPOS_LIBRO = (
    "LIBRO INDICADOR",
    "PROTOCOLO",
    "POLIZAS",
)


def normalizar_espacios(texto: str) -> str:
    """Elimina espacios iniciales/finales y reduce espacios internos a uno."""
    return " ".join(str(texto).split())


def construir_nombre_base(tipo: str, notario: str, anio: str) -> str:
    """Construye el nombre base: TIPO DE LIBRO + NOTARIO + AÑO."""
    componentes = (
        normalizar_espacios(tipo),
        normalizar_espacios(notario),
        normalizar_espacios(anio),
    )
    return " ".join(componente for componente in componentes if componente)


def obtener_nombre_base(ruta_excel) -> str:
    """Devuelve el nombre base real de una temporada a partir de su Excel."""
    return Path(ruta_excel).stem


def extraer_datos_desde_nombre(ruta_excel):
    """Extrae tipo, notario y año sin alterar el nombre existente.

    Admite la nomenclatura actual con espacios y la nomenclatura histórica
    basada en guiones bajos para poder continuar temporadas antiguas.
    """
    nombre = obtener_nombre_base(ruta_excel)
    nombre_normalizado = normalizar_espacios(nombre.replace("_", " "))
    partes = nombre_normalizado.split(" ")

    if len(partes) < 3:
        return None

    anio = partes[-1]
    if not anio.isdigit():
        return None

    cuerpo = " ".join(partes[:-1])
    cuerpo_mayusculas = cuerpo.upper()

    for tipo in TIPOS_LIBRO:
        if cuerpo_mayusculas == tipo:
            return None
        prefijo = f"{tipo} "
        if cuerpo_mayusculas.startswith(prefijo):
            notario = normalizar_espacios(cuerpo[len(prefijo):])
            if not notario:
                return None
            return tipo, notario, anio

    tipo = partes[0]
    notario = normalizar_espacios(" ".join(partes[1:-1]))
    if not notario:
        return None

    return tipo, notario, anio
