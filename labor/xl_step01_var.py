"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

"""
import os
import time

import openpyxl
import pandas as pd
from labor import Runnable
from xl_macro.dataframe_utils import save_dataframe_as
from xl_macro.langchain_xl_developer import request_doc, request_dev, PROMPT_MODEL_DOC, PROMPT_MODEL_CODE
from xl_macro.xl_macro_parser import extract_code_chunks
from xl_macro.xl_macro_reader import read_vba_macros_and_cls, read_named_ranges


class Step01(Runnable):

    def run(self):
        print(os.path.abspath("assets/input/Tarifrechner_KLV.xlsm"))
        macros = read_vba_macros_and_cls("assets/input/Tarifrechner_KLV.xlsm")
        named_ranges = read_named_ranges("assets/input/Tarifrechner_KLV.xlsm")
        extract_columns = [
            "meaning",
            "params",
            "code",
            "line_start",
            "line_number",
            "local_used",

        ]
        ws_column_types = {
            "meaning": "string",
            "params": "string",
            "code": "string",
            "line_start": "int",
            "line_number": "int",
            "local_used": "object",
            "doc_block": "string",
            "signatur": "string",
            "py_block": "string",
            "model_doc": "string",
            "doc_duration": "int",
            "model_code": "string",
            "code_duration": "int",
        }

        all_df = pd.DataFrame([], columns=ws_column_types.keys())
        all_df = all_df.astype(ws_column_types)
        for key, value in macros.items():
            if not key.endswith(".bas"):
                continue
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print(key)
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print(value)
            chunks, used = extract_code_chunks(value, named_ranges)
            df = pd.DataFrame(chunks, columns=extract_columns)
            df["doc_block"] = ""
            df["py_block"] = ""
            df["signatur"] = ""
            df["model_doc"] = ""
            df["doc_duration"] = -1
            df["model_code"] = ""
            df["code_duration"] = -1
            df = df.astype(ws_column_types)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Chunks")
            for row in df.itertuples(index=True):
                label=row.meaning
                params=row.params
                code=row.code
                von=row.line_start
                bis=row.line_number
                local_used=row.local_used
                start = time.time()
                print(f">>>\n{label} ({params}) [{von}-{bis}]=== \n{code}\n=== local usage: {local_used}")
                if label == "++Attribute++":
                    continue
                print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response doc:")
                doc_block = request_doc(label=label, code=code, full_code=value, names=used)
                print(doc_block)
                end = time.time()
                df.at[row.Index, "doc_block"] = doc_block
                df.at[row.Index, "doc_duration"] = int((end - start) * 1000)
                df.at[row.Index, "model_doc"] = PROMPT_MODEL_DOC
                if label.startswith("++"):
                    start = end
                    print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response code:")
                    py_block = request_dev(label=label, code=code, doc_block=doc_block, var_code_py = '', sign_py = [], names=used)
                    print(py_block)
                    print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end response")
                    end = time.time()
                    df.at[row.Index, "code_duration"] = int((end - start) * 1000)
                    df.at[row.Index, "model_code"] = PROMPT_MODEL_CODE
                    df.at[row.Index, "py_block"] = py_block
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~ macro usage:")
            for key, value in used.items():
                print(key, "=", value)

            all_df = pd.concat([all_df, df], ignore_index=True)

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        for key, value in named_ranges.items():
            print(key,"=",value)
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

        save_dataframe_as(all_df, "assets/output/xl_step01_var")
        all_df.to_excel("assets/output/xl_step01_var.xlsx", index=False, engine="openpyxl")
        print("Saved.")


if __name__ == "__main__":
    print("Steop01: Extract VBA macros and generate Python code snippets for vars.")
    step = Step01()
    step.run()