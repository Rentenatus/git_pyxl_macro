"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import time
from labor import Runnable
from xl_macro.dataframe_utils import save_dataframe_as, load_dataframe
from xl_macro.langchain_xl_developer import request_dev, PROMPT_MODEL_CODE


class Step03(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 03: Generate Python code snippets for methods.")

    def run(self):
        all_df = load_dataframe("assets/output/xl_step02_sign")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        py_code_start = ""
        py_doc_start = ""
        sign_dict = {}
        for idx, row in all_df.iterrows():
            meaning = row.meaning
            if meaning.startswith("++"):
                py_block = row.py_block
                doc_block = row.doc_block
                py_code_start  = py_code_start + py_block
                py_doc_start = py_doc_start + doc_block
            else:
                signatur = row.signatur
                sign_dict[meaning] = signatur

        for idx, row in all_df.iterrows():
            meaning = row.meaning
            params = row.params
            if meaning.startswith("++"):
                continue
            print(idx,":  ",meaning,"(",params,")")
            code = row.code
            used = row.local_used
            doc_block = row.doc_block
            start = time.time()
            calls = find_calls_in_code(meaning, code, sign_dict)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response code:")
            py_block = request_dev(label=meaning, code=code, doc_block=doc_block, var_code_py=py_code_start, sign_py = calls, names=used)
            end = time.time()
            all_df.at[idx, "code_duration"] = int((end - start) * 1000)
            all_df.at[idx, "model_code"] = PROMPT_MODEL_CODE
            all_df.at[idx, "py_block"] = py_block

            print(py_block)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end response")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        save_dataframe_as(all_df, "assets/output/xl_step03_code")
        all_df.to_excel("assets/output/xl_step03_code.xlsx", index=False, engine="openpyxl")
        print("Saved.")

def find_calls_in_code(omitte: str, code: str, sign_dict: dict) -> list:
    calls = []
    for sign in sign_dict.keys():
        if sign == omitte:
            continue
        if sign in code:
            calls.append(sign_dict[sign])
    return calls

if __name__ == "__main__":
    step = Step03()
    step.run()