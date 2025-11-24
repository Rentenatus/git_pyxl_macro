"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import pandas as pd
from labor import Runnable
from xl_macro.dataframe_utils import load_dataframe
from xl_macro.langchain_xl_developer import CELL_VALUE
from xl_macro.py_code_utils import code_extract, clean_import


class Step04(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 04: Rcombine the code.")

    def run(self):
        all_df = load_dataframe("assets/output/xl_step03_code")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        py_code_import = ""
        py_code_vars = ""
        py_code_methods = ""
        for idx, row in all_df.iterrows():
            meaning = row.meaning
            params = row.params
            if meaning == "++Attribute++":
                continue

            py_extract = code_extract(row.py_block)
            py_imports, py_code = clean_import(py_extract)
            if meaning.startswith("++"):
                py_code_vars = py_code_vars + "\n\n" + py_code
            else:
                print(idx,":  ",meaning,"(",params,")")
                if row.py_block is None or pd.isna(row.py_block):
                    print("Warning: Missing code block for ", meaning, "(", params, ")")
                    continue
                py_code_methods = py_code_methods + "\n\n" + py_code
            py_code_import = py_code_import + "\n" + py_imports


        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        text = py_code_import + "\n\n" + CELL_VALUE + "\n\n" + py_code_vars + "\n\n" + py_code_methods
        text = text.replace("\n\n\n\n\n", "\n\n")
        text = text.replace("\n\n\n\n", "\n\n")
        text = text.replace("\n\n\n", "\n\n")
        text = text.replace("\n\n\n", "\n\n")
        with open("assets/output/xl_recombined.py", "w", encoding="utf-8") as f:
            f.write(text)


if __name__ == "__main__":
    step = Step04()
    step.run()