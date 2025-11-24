"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

import unittest
import sys
from xl_macro.py_code_utils import code_extract, clean_import


class TestExcelMacro(unittest.TestCase):

    def setUp(self):
        pass

    def test_extracts_code_block_correctly(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        text = """Some text before
        ```python
def hello_world():
    print('Hello, World!')
        ```
        Some text after"""
        result = code_extract(text)
        print(result)
        assert result == "def hello_world():\n    print('Hello, World!')"
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_returns_empty_string_when_no_code_block(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        text = "This is a string without any code block."
        result = code_extract(text)
        print(result)
        self.assertEqual(result,"This is a string without any code block.")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_handles_multiple_code_blocks_and_returns_first(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        text = """```python
def first_function(self):
    pass
        ```
        ```python
def second_function(self):
    pass
        ```"""
        result = code_extract(text)
        print(result)
        self.assertEqual(result, 'def first_function(self):\n    pass')
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_handles_empty_code_block(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        text = "```python\n```"
        result = code_extract(text)
        print(result)
        self.assertEqual(result,"")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_separates_imports_and_code_correctly(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        source = """import os
import sys

def hello_world():
    print("Hello, World!")"""
        imports, code = clean_import(source)
        print("imports:::", imports)
        print("code:::", code)
        self.assertEqual(imports, "import os\nimport sys")
        self.assertEqual(code, "def hello_world():\n    print(\"Hello, World!\")")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_skips_specific_imports(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        source = """from excel_globals import something
from excel_math import another_thing
import os

def hello_world():
    print("Hello, World!")"""
        imports, code = clean_import(source)
        print("imports:::", imports)
        print("code:::", code)
        self.assertEqual(imports, "import os")
        self.assertEqual(code, "def hello_world():\n    print(\"Hello, World!\")")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_handles_empty_source(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        source = ""
        imports, code = clean_import(source)
        print("imports:::", imports)
        print("code:::", code)
        self.assertEqual(imports, "")
        self.assertEqual(code, "")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_handles_source_with_only_code(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        source = """def hello_world():
    print("Hello, World!")"""
        imports, code = clean_import(source)
        print("imports:::", imports)
        print("code:::", code)
        self.assertEqual(imports, "")
        self.assertEqual(code, "def hello_world():\n    print(\"Hello, World!\")")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    def test_handles_source_with_only_imports(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        source = """import os
import sys"""
        imports, code = clean_import(source)
        print("imports:::", imports)
        print("code:::", code)
        self.assertEqual(imports, "import os\nimport sys")
        self.assertEqual(code, "")
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
