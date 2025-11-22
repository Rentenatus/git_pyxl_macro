"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

"""

import time
from labor import Runnable
from xl_macro.dataframe_utils import save_dataframe_as, load_dataframe
from xl_macro.langchain_xl_developer import request_sign, PROMPT_MODEL_SIGN


class Step02(Runnable):

    def __init__(self):
        super().__init__()
        print("Step 02: Generate Python code signatures for methods.")

    def run(self):
        all_df = load_dataframe("assets/output/xl_step01_var")
        all_df["signatur"] = ""

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        py_code_start = ""
        py_doc_start = ""
        for idx, row in all_df.iterrows():
            meaning = row.meaning

            if meaning.startswith("++"):
                py_block = row.py_block
                doc_block = row.doc_block
                py_code_start  = py_code_start + py_block
                py_doc_start = py_doc_start + doc_block

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
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response signatur:")
            py_block = request_sign(label=meaning, code=code, doc_block=doc_block, var_code_py=py_code_start, names=used)
            end = time.time()
            all_df.at[idx, "code_duration"] = int((end - start) * 1000)
            all_df.at[idx, "model_code"] = PROMPT_MODEL_SIGN
            all_df.at[idx, "signatur"] = py_block

            print(py_block)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end response")

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        save_dataframe_as(all_df, "assets/output/xl_step02_sign")
        all_df.to_excel("assets/output/xl_step02_sign.xlsx", index=False, engine="openpyxl")
        print("Saved.")


if __name__ == "__main__":
    step = Step02()
    step.run()