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
from xl_macro.langchain_xl_developer import PROMPT_MODEL_CODE, request_dev_fkt
from xl_macro.py_code_utils import extract_cell_formulas
from xl_macro.xl_macro_reader import read_named_ranges


class Step04(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 05: Extract functions from the cells of the tables.")

    def run(self):
        xlsm_path = "assets/input/Tarifrechner_KLV.xlsm"
        print(os.path.abspath(xlsm_path))
        named_ranges = read_named_ranges("assets/input/Tarifrechner_KLV.xlsm")

        # Lade dein bestehendes DataFrame
        all_df = load_dataframe("assets/output/xl_step03_code")
        sign_dict = {}
        for idx, row in all_df.iterrows():
            meaning = row.meaning
            if meaning.startswith("++"):
                continue
            signatur = row.signatur
            if pd.notna(signatur): sign_dict[meaning] = signatur

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        # Sammle Formeln
        formulas = extract_cell_formulas(xlsm_path, named_ranges.keys(), sign_dict.keys())
        print(f"Gefundene Formeln: {len(formulas)}")

        fkt_column_types = {
            "sheet_title": "string",
            "coord": "string",
            "value_type": "string",
            "fkt_name": "string",
            "fkt_code": "string",
            "used_names": "object",
            "used_meanings": "object",

        }
        fkt_df = pd.DataFrame(formulas.values(), columns=fkt_column_types.keys())
        fkt_df = fkt_df.astype(fkt_column_types)
        fkt_df["py_fkt"] = ""
        fkt_df["model_code"] = ""
        fkt_df["code_duration"] = -1

        for idx, row in fkt_df.iterrows():
            cell_ref = row.coord
            formel_code = row.fkt_code
            method_name = row.fkt_name
            names = row.used_names
            used_meanings = row.used_meanings
            print(cell_ref, "->", row.value_type, ":", method_name, "using", len(used_meanings), "meanings.")
            start = time.time()
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response function:")
            response = request_dev_fkt(cell_ref, formel_code, method_name,names, used_meanings)
            print(response)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end response")
            end = time.time()
            fkt_df.at[idx, "code_duration"] = int((end - start) * 1000)
            fkt_df.at[idx, "model_code"] = PROMPT_MODEL_CODE
            fkt_df.at[idx, "py_fkt"] = response

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        save_dataframe_as(fkt_df, "assets/output/xl_step04_fkt")
        fkt_df.to_excel("assets/output/xl_step04_fkt.xlsx", index=False, engine="openpyxl")
        print("Saved.")


if __name__ == "__main__":
    step = Step04()
    step.run()