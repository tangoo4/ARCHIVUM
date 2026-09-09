"""Validaciones y normalización compartidas de datos funcionales de ARCHIVUM."""


def normalizar_anio(valor):
    texto = str(valor).strip()
    if not texto.isdigit():
        raise ValueError("El año debe ser un número entero.")
    return int(texto)


def normalizar_medida(valor, maximo=10):
    texto = str(valor).strip()
    if texto == "?":
        return "?"
    try:
        numero = float(texto.replace(",", "."))
    except (TypeError, ValueError):
        raise ValueError("La medida debe ser un número o ?.")
    if numero <= 0 or numero > maximo:
        raise ValueError(f"La medida debe estar entre 0 y {maximo}.")
    return round(numero, 1)


def formatear_medida(valor):
    if valor == "?":
        return "?"
    if valor in (None, ""):
        return ""
    try:
        return f"{float(str(valor).replace(',', '.')):.1f}".replace(".", ",")
    except (TypeError, ValueError):
        return str(valor)
