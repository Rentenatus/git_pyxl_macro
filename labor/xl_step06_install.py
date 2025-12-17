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


class Step06(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 06: Install python code.")

    def run(self):
        py_file = "assets/output/xl_recombined.py"
        target_file = "preprod/xl_code.py"

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        # Stelle sicher, dass Zielverzeichnis existiert
        os.makedirs(os.path.dirname(target_file), exist_ok=True)

        # Kopiere Datei
        try:
            shutil.copy(py_file, target_file)
            print(f"Copied {py_file} -> {target_file}")
        except FileNotFoundError:
            print(f"Source file not found: {py_file}")
        except Exception as e:
            print(f"Error while copying: {e}")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print("Installed.")


if __name__ == "__main__":
    step = Step06()
    step.run()
