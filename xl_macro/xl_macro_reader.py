# -*- coding: utf-8 -*-
"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

from oletools.olevba import VBA_Parser
from openpyxl import load_workbook
import zipfile
import re

def has_vba(file_path: str) -> bool:
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            return any(f.startswith('xl/vbaProject') for f in z.namelist())
    except Exception:
        return False

def read_vba_macros_and_cls(file_path: str) -> dict[str, str]:
    macros = {}
    if not has_vba(file_path):
        return macros
    parser = VBA_Parser(file_path)
    if parser.detect_vba_macros():
        for (_, _, macro_name, macro_code) in parser.extract_macros():
            macros[macro_name] = macro_code
    return macros

def read_named_ranges(file_path: str) -> dict[str, str]:
    wb = load_workbook(file_path, data_only=True, keep_links=False)
    names = wb.defined_names
    ret = {}
    for nn in names:
        ret[nn] = names.get(nn).value
    return ret



def extract_used_names_from_vba(code: str, defined_names: list[str]) -> list[str]:
    """
    Extrahiert verwendete Namen aus VBA-Code, basierend auf direkten Gänsefüßchen-Zugriffen.
    Kommentare werden ignoriert.
    """
    used = set()
    name_set = set(n.lower() for n in defined_names)

    for line in code.splitlines():
        line = line.strip()

        # Zustand: Kommentar → überspringen
        if line.startswith("'") or line.lower().startswith("rem"):
            continue

        # Zustand: Code → Strings extrahieren
        string_literals = re.findall(r'"([^"]+)"', line)
        for literal in string_literals:
            if literal.lower() in name_set:
                used.add(literal)

    return sorted(used)
