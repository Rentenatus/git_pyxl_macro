# -*- coding: utf-8 -*-
import re

TOKEN_REGEX = re.compile(
    r"""
    (?P<STRING>"(?:[^"\\]|\\.)*")        |
    (?P<COMMENT>'[^\n\r]*)               |
    (?P<IDENTIFIER>[A-Za-z_][A-Za-z0-9_]*)|
    (?P<NUMBER>\d+(?:\.\d+)?)            |
    (?P<LPAREN>\()                       |
    (?P<RPAREN>\))                       |
    (?P<LBRACKET>\[)                     |
    (?P<RBRACKET>\])                     |
    (?P<SPACE>[ \t]+)                    |
    (?P<NEWLINE>\r?\n)                   |
    (?P<OTHER>[<>=.,!§$%&/\\?\{\}:+\-*^|~#@])
    """,
    re.VERBOSE
)


def tokenize_vba(code: str):
    pos = 0
    tokens = []
    while pos < len(code):
        match = TOKEN_REGEX.match(code, pos)
        if not match:
            raise SyntaxError(f"Unrecognized character at position {pos}: {code[pos]}")
        kind = match.lastgroup
        value = match.group()
        tokens.append((kind, value))
        pos = match.end()
    return tokens

def extract_code_chunks(code: str, named_ranges: dict[str, str] = {}):
    tokens = tokenize_vba(code)
    return reconstruct_code(tokens, named_ranges)


def reconstruct_code(tokens, named_ranges: dict[str, str] = {}):
    chunks = []
    row = ""
    code = ""
    meaning = "++Start++"
    reading_param = False
    waiting_for_param = False
    waiting_for_name = False
    waiting_end_for = False
    procedure_type = ""
    line_number = 1
    line_start = line_number
    used={}
    local_used={}
    params = []
    for kind, value in tokens:
        if kind == "NEWLINE":
            line_number += 1
            code += row + value
            row = ""
            continue
        elif kind == "LPAREN":
            row += value
            reading_param = waiting_for_param
            waiting_for_param = False
            continue
        elif kind  == "RPAREN":
            row += value
            reading_param = waiting_for_param = False
            continue

        if reading_param:
            if kind != "SPACE": params.append(value)
            row += value
            continue

        if kind == "IDENTIFIER":
            upper = value.upper()
            if waiting_for_name:
                row += value
                meaning = value
                waiting_for_name = False
                waiting_for_param = True
                params = []
            elif waiting_end_for and procedure_type == upper:
                code += row + value
                row = ""
                chunks.append((meaning, params, code, line_start, line_number, local_used))
                code = ""
                local_used = {}
                meaning = "++Next++"
                procedure_type = ""
                line_start = line_number + 1
                reading_param = waiting_for_param = False
                params = []
            elif upper == "END":
                row += value
                waiting_end_for = True
                reading_param = waiting_for_param = False
            elif upper in ("FUNCTION", "PROPERTY", "SUB"):
                if code.strip():
                    chunks.append((meaning, params, code, line_start, line_number-1, local_used))
                code = ""
                local_used = {}
                row += value
                meaning = f"++Unknown_{value}++"
                waiting_for_name = True
                procedure_type = upper
                waiting_for_param = False
                params = []
                line_start = line_number
            elif upper in ("DIM", "CONST", "DECLARE", "LET") and not procedure_type:
                if meaning != "++Declaration++":
                    if code.strip():
                        chunks.append((meaning, params, code, line_start, line_number-1, local_used))
                    code = ""
                    local_used = {}
                row += value
                meaning = "++Declaration++"
                line_start = line_number
            elif (upper == 'ATTRIBUTE') and not row.strip():
                if meaning != "++Attribute++":
                    if code.strip():
                        chunks.append((meaning, params, code, line_start, line_number-1, local_used))
                    code = ""
                    local_used = {}
                row += value
                meaning = "++Attribute++"
                line_start = line_number
            else:
                if (value in named_ranges.keys()) and not (value in params):
                    local_used[value] = named_ranges[value]
                    used[value] = named_ranges[value]
                row += value
        else:
            row += value
    if row.strip():
        code += row
    if code.strip():
        chunks.append((meaning, params, code, line_start, line_number, local_used))

    return chunks, used



