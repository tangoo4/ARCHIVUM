"""Operaciones de lectura y modificación de tomos ya guardados."""

import copy
from pathlib import Path

from openpyxl import load_workbook

from core.validaciones import normalizar_anio, normalizar_medida
from excel.lectura import leer_tomos_temporada
from excel.estilos import aplicar_estado_fila
from config import MAX_MEDIDA


FILA_INICIO_DATOS = 2
COLUMNAS_TOMO = "ABCDEFGH"

def leer_tomos(ruta_excel):
    """Lee los tomos usando el criterio común de Medido."""
    tomos = leer_tomos_temporada(ruta_excel)
    for tomo in tomos:
        tomo["fila"] = tomo.pop("fila_excel")
    return tomos


def comprobar_continuidad(tomos, indice, matriz_inicio, matriz_final):
    """Devuelve avisos si la corrección rompe la continuidad con otro tomo."""
    avisos = []
    if indice > 0:
        try:
            esperado = int(tomos[indice - 1]["matriz_final"]) + 1
            if matriz_inicio != esperado:
                avisos.append(f"La matriz inicial esperada según el tomo anterior es {esperado}.")
        except (TypeError, ValueError):
            pass

    if indice < len(tomos) - 1:
        try:
            esperado_siguiente = matriz_final + 1
            inicio_siguiente = int(tomos[indice + 1]["matriz_inicio"])
            if inicio_siguiente != esperado_siguiente:
                avisos.append(
                    f"El tomo siguiente empieza en {inicio_siguiente}; debería empezar en {esperado_siguiente}."
                )
        except (TypeError, ValueError):
            pass
    return avisos


def modificar_tomo(ruta_excel, tomo, datos, medida_estandar):
    """Actualiza A:H en la fila del tomo sin alterar foliado ni columnas auxiliares."""
    ruta_excel = Path(ruta_excel)
    wb = load_workbook(ruta_excel, data_only=False)
    ws = wb.active

    fila_objetivo = None
    for fila in range(FILA_INICIO_DATOS, ws.max_row + 1):
        try:
            if int(ws[f"A{fila}"].value) == int(tomo):
                fila_objetivo = fila
                break
        except (TypeError, ValueError):
            continue

    if fila_objetivo is None:
        raise ValueError(f"No existe el tomo {tomo} en la temporada.")

    valores = (
        int(tomo),
        normalizar_anio(datos["anio"]),
        int(datos["matriz_inicio"]),
        datos["fecha_inicio"],
        int(datos["matriz_final"]),
        datos["fecha_final"],
        normalizar_medida(datos["medida"], MAX_MEDIDA),
        datos["observaciones"],
    )
    for columna, valor in zip(COLUMNAS_TOMO, valores):
        ws[f"{columna}{fila_objetivo}"] = valor

    medida = datos["medida"]
    if medida != "?":
        ws[f"G{fila_objetivo}"].number_format = "0.0"
    aplicar_estado_fila(ws, fila_objetivo, medida, COLUMNAS_TOMO)

    wb.save(ruta_excel)

