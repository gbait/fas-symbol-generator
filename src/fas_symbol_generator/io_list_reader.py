"""
Lee una hoja de lista de I/O (formato ST03/AZR estándar de ALVEA FAS3G)
y devuelve la lista de equipos únicos que contiene, sin duplicar instancias
repetidas (una fila por atributo, pero un equipo puede tener varias filas).
"""

import openpyxl


# Nombres de columna tal como aparecen en la fila de cabecera real (fila 4
# en los ficheros ST03 que hemos visto). Si tu hoja usa otros nombres,
# ajusta este diccionario y el resto del código no cambia.
COLUMNAS = {
    "Eqpt_Code": None,
    "Eqpt_Name": None,
    "Eqpt_Description": None,
    "Eqpt_Identifier": None,
}


def _localizar_columnas(hoja, fila_cabecera):
    """Busca en la fila de cabecera el número de columna de cada campo
    que necesitamos, para no depender de que estén siempre en el mismo
    orden exacto."""
    columnas = {}
    for columna in range(1, hoja.max_column + 1):
        valor = hoja.cell(row=fila_cabecera, column=columna).value
        if valor in COLUMNAS:
            columnas[valor] = columna
    faltantes = [c for c in COLUMNAS if c not in columnas]
    if faltantes:
        raise ValueError(
            f"No se han encontrado estas columnas en la fila {fila_cabecera}: {faltantes}"
        )
    return columnas


def leer_equipos(ruta_excel, nombre_hoja, fila_cabecera=4, fila_inicio_datos=5):
    """
    Devuelve una lista de diccionarios, uno por cada equipo ÚNICO
    (por Eqpt_Name), con las claves: codigo, nombre, descripcion, entity_id.

    ruta_excel: ruta al fichero .xlsx / .xlsm
    nombre_hoja: nombre exacto de la pestaña, p.ej. "IO List ST03 - RS"
    fila_cabecera: fila donde están los títulos de columna (por defecto 4)
    fila_inicio_datos: primera fila con datos reales (por defecto 5)
    """
    libro = openpyxl.load_workbook(ruta_excel, data_only=True)
    hoja = libro[nombre_hoja]
    columnas = _localizar_columnas(hoja, fila_cabecera)

    equipos = []
    nombres_vistos = set()

    for fila in range(fila_inicio_datos, hoja.max_row + 1):
        nombre = hoja.cell(row=fila, column=columnas["Eqpt_Name"]).value
        codigo = hoja.cell(row=fila, column=columnas["Eqpt_Code"]).value
        entity_id = hoja.cell(row=fila, column=columnas["Eqpt_Identifier"]).value

        if not nombre or not entity_id:
            continue  # fila vacía o incompleta, se ignora
        if nombre in nombres_vistos:
            continue  # ya lo tenemos (el mismo equipo aparece en varias filas)

        nombres_vistos.add(nombre)
        equipos.append({
            "codigo": codigo,
            "nombre": nombre,
            "descripcion": hoja.cell(row=fila, column=columnas["Eqpt_Description"]).value,
            "entity_id": entity_id,
        })

    return equipos


if __name__ == "__main__":
    # Prueba rápida manual: python io_list_reader.py
    equipos = leer_equipos("../../data/sample_io_list.xlsx", "IO List Sample")
    for e in equipos:
        print(e)
