# -*- coding: utf-8 -*-
"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""


import re
import os
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula


def extract_cell_formulas(xlsm_path: str) -> dict:
    """
    Läuft über alle Blätter und Zellen einer Excel-Datei (.xlsm)
    und sammelt alle Formeln in einem Dict.
    Key = Zellenkoordinate (z.B. 'A1'), Value = Formelstring.
    """
    wb = openpyxl.load_workbook(xlsm_path, data_only=False)  # data_only=False => Formeln statt Werte
    formulas = {}

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                fkt_code = None
                value_type = ""
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    fkt_code = str(cell.value)
                    value_type = "str"
                elif cell.value and isinstance(cell.value, ArrayFormula):
                    fkt_code = cell.value.text
                    value_type = "ArrayFormula"

                if fkt_code is None:
                    continue

                sheet_title = sheet.title
                coord = f"{sheet.title}!{cell.coordinate}"
                fkt_name = f"fkt_{sheet.title}_{cell.coordinate}".lower()
                formulas[coord] = (sheet_title, coord, value_type, fkt_name, fkt_code)

    return formulas

def code_extract(text: str) -> str:
    """
    Extrahiert den Python-Code aus einem String, der mit ```python beginnt
    und mit ``` endet. Gibt den reinen Code zurück oder den kompletten text.
    """
    if "```python" not in text:
        return text
    # Regex: sucht nach ```python ... ```
    match = re.search(r"```python\s*(.*?)\s*```", text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""


def clean_import(source: str) -> tuple[str, str]:
    """
    Trennt den gegebenen Python-Quelltext in Import-Teil und Code-Teil.
    Überspringt bestimmte Imports (excel_globals, excel_math).

    Args:
        source: kompletter Python-Quelltext als String

    Returns:
        (imports_str, code_str)
    """
    import_lines = []
    code_lines = []

    for line in source.splitlines():
        stripped = line.strip()
        if stripped.startswith("import ") or stripped.startswith("from "):
            # überspringe bestimmte Imports
            if stripped.startswith("from excel_globals import") or stripped.startswith("from excel_math import"):
                continue
            import_lines.append(line)
        else:
            code_lines.append(line)

    imports_str = "\n".join(import_lines).strip()
    code_str = "\n".join(code_lines).strip()
    return imports_str, code_str
