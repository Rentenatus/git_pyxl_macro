"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import unittest

from xl_macro.langchain_xl_developer import request_doc, request_dev
from xl_macro.xl_macro_parser import extract_code_chunks
from xl_macro.xl_macro_reader import read_vba_macros_and_cls, read_named_ranges


class TestExcelMacro(unittest.TestCase):

    def setUp(self):
        pass

    def test_compare_and_update_expected(self):
        macros = read_vba_macros_and_cls("test/assets/input/Tarifrechner_KLV.xlsm")
        named_ranges = read_named_ranges("test/assets/input/Tarifrechner_KLV.xlsm")
        for key, value in macros.items():
            if not key.endswith(".bas"):
                continue
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print(key)
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print(value)
            chunks, used = extract_code_chunks(value, named_ranges)
            print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Chunks")
            for label, params, code, von, bis, local_used in chunks:
                print(f">>>\n{label} ({params}) [{von}-{bis}]=== \n{code}\n=== local usage: {local_used}")
                if label == "++Attribute++":
                    continue
                if label.startswith("++"):
                    print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response 1:")
                    doc_block = request_doc(label=label, code=code, full_code=value, names=used)
                    print(doc_block)
                    print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ response 2:")
                    print(request_dev(label=label, code=code, doc_block=doc_block, var_code_py="", sign_py = [], names=used))
                    print("#######~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end response")

            print("~~~~~~~~~~~~~~~~~~~~~~~~~~ macro usage:")
            for key, value in used.items():
                print(key, "=", value)
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        for key, value in named_ranges.items():
            print(key,"=",value)
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
