"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>
 
"""

import unittest

from xl_macro.xl_macro_parser import  extract_code_chunks


class TestExcelMacro(unittest.TestCase):

    def setUp(self):
        pass

    def test_simple_parse(self):
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Input:")

        test_code = """
Dim a As Integer
a = Zins
Sub Test(n As Integer)
    Dim x As Integer
    x = zw
    MsgBox "Hello"+n
End Sub

Dim a As Integer

"""
        print(test_code)

        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Output:")
        chunks, _ = extract_code_chunks(test_code, {
            "Zins":"Kalkulation!$E$4",
            "n":"Kalkulation!$B$6",
            "zw":"Kalkulation!$B$9"
        })
        print(chunks)

        expected = [('++Declaration++', [], 'Dim a As Integer\na = Zins\n', 2, 3, {'Zins': 'Kalkulation!$E$4'}),
                    ('Test', ['n', 'As', 'Integer'],
                        'Sub Test(n As Integer)\n    Dim x As Integer\n    x = zw\n    MsgBox "Hello"+n\nEnd Sub',
                        4, 8, {'zw': 'Kalkulation!$B$9'}),
                    ('++Declaration++', [], 'Dim a As Integer\n\n', 10, 12, {})]



        self.assertEqual(chunks, expected)