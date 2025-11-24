# -*- coding: utf-8 -*-
"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""


import re

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
