# -*- coding: utf-8 -*-
"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

from langchain_ollama import ChatOllama
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from langchain_core.prompts import PromptTemplate

CELL_NAME_VALUE = """from openpyxl import Workbook
from openpyxl.utils import range_boundaries

xl_workbook: Workbook # The Excel workbook object
xl_names: dict [str, str] # The dictionary of named ranges

def get_excel_global(key: str):
    \"\"\"
    Retrieves the value from the Excel workbook based on a named range key.
    
    Parameters:
    - key (str): The name of the range to look up.
    
    Returns:
    - The value of the cell referenced by the named range, or None if not found.
    \"\"\"
    if key not in xl_names:
        raise KeyError(f"Key '{key}' not found in xl_names.")

    ref = xl_names[key]  # e.g., 'Kalkulation!$E$6'
    if '!' not in ref:
        raise ValueError(f"Invalid reference format: '{ref}'")

    sheet_name, cell_ref = ref.split('!')
    sheet_name = sheet_name.strip("'")  # remove quotes if present
    cell_ref = cell_ref.replace('$', '')  # remove dollar signs

    if sheet_name not in xl_workbook.sheetnames:
        raise ValueError(f"Worksheet '{sheet_name}' not found in workbook.")

    sheet = xl_workbook[sheet_name]
    cell = sheet[cell_ref]
    return cell.value"""


CELL_VALUE = """
def get_cell_value(ref: str):
    try:
        sheet_name, cell_ref = ref.split("!", 1)
        return get_cell_value2(sheet_name, cell_ref)
    except ValueError:
        raise ValueError(f"Ungültiges Format: {ref}. Erwartet 'Sheet!Cell'.")


def get_cell_value2(sheet_name: str, cell_ref: str):
    # Funktionsname nach deinem Schema konstruieren
    func_name = f"fkt_{sheet_name.lower()}_{cell_ref.lower()}"

    # Prüfen, ob die Funktion existiert
    if func_name in globals():
        return globals()[func_name]()
    else:
        # Wert direkt aus Excel lesen
        sheet = xl_workbook[sheet_name]
        return sheet[cell_ref].value
        
"""



BASE_URL = "http://127.0.0.1:11434"

# PromptModel: Modelle für die verschiedenen Aufgaben
#############################################################################
PROMPT_MODEL_DOC = "gemma3:27b"
PROMPT_MODEL_SIGN = "gemma3:27b"
PROMPT_MODEL_CODE = "devstral-small-2:24b"

# Probiere es mit verschiedenen Modellen mal aus:
# PROMPT_MODEL_DOC = "gpt-oss:20b"
# PROMPT_MODEL_SIGN = "qwen2.5-coder:32b"
# PROMPT_MODEL_CODE = "gemma3:27b"

# Systemprompt: Rollenbeschreibung
#############################################################################

SYSTEM_PROMPT_DOC = """You are a software architect and you will document VBA code very effectively.
You are also an expert in German pensions.
For example, the following knowledge is second nature to you:

Das 3-Schichten-Modell der Altersvorsorge in deutschlad entstand im Jahr 2005 mit Einführung des Alterseinkünftegesetzes (AltEinkG) und löste das bis dahin geltende 3-Säulen-Modell durch die Darstellung aus Sicht der jeweiligen Regeln zur Besteuerung ab.

In der 1. Schicht finden sich die gesetzliche Rentenversicherung, berufsständische Versorgungswerke, die landwirtschaftlichen Alterskassen und die Beamtenversorgung.

Zur 2. Schicht der Altersvorsorge gehören staatlich geförderte Zusatzversorgung, die betriebliche Altersvorsorge und die Riester Rente.

Zur 3. Schicht der Altersvorsorge gehören alle Vorsorgeverträge, für die es in der Ansparphase keine staatliche Förderung gibt, wie private Kapitalanlagen.
"""

SYSTEM_PROMPT_DEV_DEF = """You are an expert software developer and are currently translating a piece of VBA code into Python.

Global names, if used, can be retrieved using get_excel_global(key:str).

Assume that this method '''def get_excel_global(key: str):''' is already defined in previous sections and usable now. Don't rewrite it. The code of this method is:
'''
%s
'''

The user gives you individual code pieces step by step, either a declaration or a method, and if available, a description from the requirements.

You simply rewrite this piece into Python code. All your building blocks will later be assembled into a Python file. This will create a complete program.
""" % CELL_NAME_VALUE

SYSTEM_PROMPT_DEV_VAR = """You are an expert software developer and are currently translating a piece of VBA code into Python.

Global names, if used, can be retrieved using get_excel_global(key:str).

Assume that this method '''def get_excel_global(key: str):''' is already defined in previous sections and usable now. Don't rewrite it. The code of this method is:
'''
%s
'''

The user first gives you the definition of variables.

You simply rewrite this piece into Python code. Each of your pieces is glued together in a Python file. This creates a complete program.
""" % CELL_NAME_VALUE


SYSTEM_PROMPT_DEV_FKT = """You are an expert software developer and are currently translating a piece of excel formula into Python.

Global names, if used, can be retrieved using get_excel_global(key:str).

Assume that this method '''def get_excel_global(key: str)''' is already defined in previous sections and usable now. Don't rewrite it. The code of this method is:
'''
%s
'''

Cell values can be retrieved using the function get_cell_value(ref:str).

Assume that this method '''def get_cell_value(ref: str)''' is already defined in previous sections and usable now. Don't rewrite it. The code of this method is:
'''
%s
'''

The user first gives you the definition of excel function.

You simply rewrite this piece into Python code. All blocks will later be assembled into a Python file.
""" % (CELL_NAME_VALUE, CELL_VALUE)

# PromptTemplate für den Code
#############################################################################
READ_PROMPT_TEMPLATE = """First, just read the following code:
'''
{full_code}
'''
"""

UNDERSTOOD_PROMPT_TEMPLATE="""I have read the code. I understand it."""

# PromptTemplate für den Userprompt
#############################################################################
USER_PROMPT_TEMPLATE_START_PY = """
Before we begin, I'd like to give you the existing code in python:
'''
{py_code}
'''
"""

#  PromptTemplate für den Userprompt, um die Dokumentation der Methode zu erzeugen.
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DOC_DEF = """
Document this method and only this method from a VBA code:
'''
{code}
'''

The following names are used in this method:
{names_block}

Limit yourself to this method, its parameters, its functionality, its meaning, and its relationship to the macro.
Here is the method you should document again:
'''
{code}
'''
"""

#  PromptTemplate für den Userprompt, um die Dokumentation der Variablendeklaration zu erzeugen
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DOC_VAR = """
Document this variables and only this variables from a VBA code:
'''
{code}
'''

Limit yourself to this variable.
Here is the code you should document again:
'''
{code}
'''
"""

#  PromptTemplate für den Userprompt für das Development, um die Signatur der Methode zu erzeugen.
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DEV_SIGN = """
I have a bit of VBA code here.
{doc_block}

The following names are used in this method:
{names_block}

The Python code already starts with this import:
'''python
from excel_globals import xl_workbook, xl_names, get_excel_global
'''

Your task is simply to generate the method signature from VBA into Python. Just that one line!. 
This line begins with 'def' followed by the method name in python style: 
An underscore precedes the originally capital letters; all letters are lowercase.
The VBA code reads:
'''
{code}
'''
"""

#  PromptTemplate für den Userprompt für das Development, um den Code der Methode zu erzeugen.
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DEV_DEF = """
I have a bit of VBA code here.
{doc_block}

The following names are used in this method:
{names_block}

{additional_instructions}

Start the python code with this import:
'''python
from excel_globals import xl_workbook, xl_names, get_excel_global
'''

You will translate it into Python code. Limit yourself to this one method. 
Method line begins with 'def' followed by the method name. 
When using global variables, remember to include the keyword "global" in the method, e.g.:
'''python
def my_function():
    global global_var = []
    # function body
'''


Please use this signature:
{signature}

The VBA code reads:
'''
{code}
'''
"""

USER_PROMPT_TEMPLATE_DEV_DEF_ADD_SIGN = """

The following methods have already been provided in earlier blocks; use 'from excel_math import ...' to retrieve them.; use them if needed:
{sign}

 """

#  PromptTemplate für den Userprompt für das Development, um Code der Variablendeklaration zu erzeugen.
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DEV_VAR = """
I have a bit of VBA code here.
{doc_block}

Start the python code with this import:
'''python
from excel_globals import xl_workbook, xl_names, get_excel_global
'''

You will translate it into Python code. Limit yourself to this variable. The VBA code reads:
'''
{code}
'''
"""

#  PromptTemplate für den Userprompt für das Development, um die Excelformel zu erzeugen.
#  ----------------------------------------------------------------------------
USER_PROMPT_TEMPLATE_DEV_FKT = """
{text_call_following_functions}
{functions_block}

The following names are used in this method:
{names_block}

You could define local variables with the following code:
'''python
{local_variables}
'''

To access cells from the formula, you use this method get_cell_value(ref: str), imported from excel_globals. 
It automatically checks whether it's a value or another formula. You can rely on it.
ref has the formatting "sheet!cell".
Example:
'''python
    value = get_cell_value("{cell_ref}")
'''

The Python code already starts with this import:
'''python
from excel_globals import *
'''

Your task is to generate the formula into Python. Just that one formula!. 
This line begins with 'def' followed by the method name '{method_name}' in python style: 
The excel formula code reads:

{formel_code}

"""

TEXT_CALL_FOLLOWING_FUNCTIONS = """
You can call the following functions in your code. They are already implemented in excel_globals:
"""

# Anmerkung: Ohne "Start the python code with this import:" wird das LLama immer Angst haben unperfekt zu sein
# und die Methode get_excel_global rezitieren.
# Da es mit einem Import beginnen soll, ist es mutiger und schreibt nur und nur den angefragten Code.
# Im trainierten Herzen des Modells steckt: "Du musst immer vollständig sein und alles liefern, was du kannst."
# Der Text "Start the python code with this import:" ist quasi eine Injection, die das Modell beschwichtigt.

def prompt_signatur(code: str, doc_block: str, var_code_py: str, names: str) -> list:
    names_block = names if names else "None"
    start_py_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_START_PY).format(
        py_code=var_code_py
    )
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DEV_SIGN).format(
        code=code, doc_block=doc_block, names_block=names_block
    )
    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DEV_DEF),
        HumanMessage(content=start_py_prompt),
        AIMessage(content=UNDERSTOOD_PROMPT_TEMPLATE),
        HumanMessage(content=user_prompt)
    ]
    return messages

def prompt_dev_def(code: str, doc_block: str, var_code_py: str, sign_py, own_sign: str, names: str) -> list:
    names_block = names if names else "None"
    additional_instructions = ""
    if sign_py:
        additional_instructions = USER_PROMPT_TEMPLATE_DEV_DEF_ADD_SIGN.format(
            sign=sign_py
        )
    start_py_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_START_PY).format(
        py_code=var_code_py
    )
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DEV_DEF).format(
        code=code, doc_block=doc_block, names_block=names_block, additional_instructions=additional_instructions,
        signature=own_sign
    )
    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DEV_DEF),
        HumanMessage(content=start_py_prompt),
        AIMessage(content=UNDERSTOOD_PROMPT_TEMPLATE),
        HumanMessage(content=user_prompt)
    ]
    return messages

def prompt_dev_var(code: str, doc_block: str, names: str) -> list:
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DEV_VAR).format(
        code=code, doc_block=doc_block
    )
    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DEV_VAR),
        HumanMessage(content=user_prompt)
    ]
    return messages

def prompt_doc_def(code: str, full_code: str, names: str) -> list:
    names_block = names if names else "None"
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DOC_DEF).format(
        code=code, full_code=full_code, names_block=names_block
    )
    code_prompt = PromptTemplate.from_template(READ_PROMPT_TEMPLATE).format(
        full_code=full_code
    )
    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DOC),
        HumanMessage(content=code_prompt),
        AIMessage(content=UNDERSTOOD_PROMPT_TEMPLATE),
        HumanMessage(content=user_prompt)
    ]
    return messages

def prompt_doc_var(code: str, full_code: str, names: str) -> list:
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DOC_VAR).format(
        code=code, full_code=full_code
    )
    code_prompt = PromptTemplate.from_template(READ_PROMPT_TEMPLATE).format(
        full_code=full_code
    )
    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DOC),
        HumanMessage(content=code_prompt),
        AIMessage(content=UNDERSTOOD_PROMPT_TEMPLATE),
        HumanMessage(content=user_prompt)
    ]
    return messages

def get_response(messages: list, model: str) -> str:
    llm = ChatOllama(model=model, base_url=BASE_URL)
    response = llm.invoke(messages)
    return response.content

def request_doc(label: str, code: str, full_code: str, names: str) -> str:
    if label.startswith("++"):
        messages = prompt_doc_var(code, full_code, names)
    else:
        messages = prompt_doc_def(code, full_code, names)
    response = get_response(messages, model=PROMPT_MODEL_DOC)
    return response

def request_dev(label: str, code: str, doc_block: str, var_code_py: str, sign_py, own_sign:str, names: str) -> str:
    if label.startswith("++"):
        messages = prompt_dev_var(code, doc_block, names)
    else:
        messages = prompt_dev_def(code, doc_block, var_code_py, sign_py, own_sign, names)
    response = get_response(messages, model=PROMPT_MODEL_CODE)
    return response

def request_sign(label: str, code: str, doc_block: str, var_code_py: str, names: str) -> str:
    messages = prompt_signatur(code, doc_block, var_code_py ,names)
    response = get_response(messages, model=PROMPT_MODEL_SIGN)
    return response

def prompt_dev_fkt(cell_ref: str, formel_code: str, method_name: str,
                   names, used_py) -> list:
    names_block = names if names else "None"
    local_variables = [] if names else ["    local_variable = get_excel_global('local_variable')"]
    for name in names:
        local_variables.append(f"    {name} = get_excel_global('{name}')")
    functions_block = used_py if used_py else ""
    text_call = TEXT_CALL_FOLLOWING_FUNCTIONS if used_py else ""
    user_prompt = PromptTemplate.from_template(USER_PROMPT_TEMPLATE_DEV_FKT).format(
        text_call_following_functions = text_call,
        functions_block = functions_block,
        names_block = names_block,
        local_variables = local_variables,
        cell_ref = cell_ref,
        formel_code = formel_code,
        method_name = method_name
    )

    # Promptliste mit Konversationsverlauf
    messages = [
        SystemMessage(content=SYSTEM_PROMPT_DEV_FKT),
        HumanMessage(content=user_prompt)
    ]
    return messages

def request_dev_fkt(cell_ref: str, formel_code: str, method_name: str,
                   names, used_py) -> str:
    messages = prompt_dev_fkt(cell_ref, formel_code, method_name, names, used_py)
    response = get_response(messages, model=PROMPT_MODEL_CODE)
    return response