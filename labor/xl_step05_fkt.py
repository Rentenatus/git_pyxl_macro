"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import os
import time

import pandas as pd
from labor import Runnable
from xl_macro.dataframe_utils import save_dataframe_as, load_dataframe
from xl_macro.py_code_utils import extract_cell_formulas
from xl_macro.xl_macro_reader import read_named_ranges


class Step05(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 05: Extract functions from the cells of the tables.")

    def run(self):
        xlsm_path = "assets/input/Tarifrechner_KLV.xlsm"
        print(os.path.abspath(xlsm_path))
        named_ranges = read_named_ranges("assets/input/Tarifrechner_KLV.xlsm")

        # Lade dein bestehendes DataFrame
        all_df = load_dataframe("assets/output/xl_step03_code")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        # Sammle Formeln
        formulas = extract_cell_formulas(xlsm_path)
        print(f"Gefundene Formeln: {len(formulas)}")
        # Beispielausgabe
        for k, v in list(formulas.items()):
            print(k, "->", v)



        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        #save_dataframe_as(all_df, "assets/output/xl_step05_fkt")
        #all_df.to_excel("assets/output/xl_step05_fkt.xlsx", index=False, engine="openpyxl")
        #print("Saved.")


if __name__ == "__main__":
    step = Step05()
    step.run()