"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import os
import shutil
from labor import Runnable


class Step09(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 06: NG-Test.")

    def run(self):
        py_file = "preprod/xl_code.py"

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        # starte den externen Test
        print("Running tests...")
        result = test_ng_test(py_file)  # implementiere diese Funktion entsprechend

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print("Result:\n", result)

import subprocess

def test_ng_test(py_file: str) -> str:
    """
    Führt NGTest für die angegebene Python-Datei aus und gibt die Ausgabe zurück.
    """
    try:
        # Beispiel: ngtest <file>
        result = subprocess.run(
            ["ngtest", py_file],
            capture_output=True,
            text=True,
            check=False  # wir fangen Fehler selbst ab
        )
        output = result.stdout
        if result.stderr:
            output += "\n[stderr]\n" + result.stderr
        return output
    except FileNotFoundError:
        return "Fehler: NGTest ist nicht installiert oder nicht im PATH."
    except Exception as e:
        return f"Unerwarteter Fehler: {e}"


if __name__ == "__main__":
    step = Step09()
    step.run()
