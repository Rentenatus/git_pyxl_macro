# -*- coding: utf-8 -*-

from langchain_ollama import ChatOllama
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from langchain_core.prompts import PromptTemplate

BASE_URL = "http://127.0.0.1:11434"
PROMPT_MODEL_DOC = "gemma3:27b"
PROMPT_MODEL_SIGN = "gemma3:27b"
PROMPT_MODEL_CODE = "gemma3:27b"

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
from openpyxl import Workbook
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
    return cell.value
'''

The user gives you individual code pieces step by step, either a declaration or a method, and if available, a description from the requirements.

You simply rewrite this piece into Python code. All your building blocks will later be assembled into a Python file. This will create a complete program.
"""

SYSTEM_PROMPT_DEV_VAR = """You are an expert software developer and are currently translating a piece of VBA code into Python.

Global names, if used, can be retrieved using get_excel_global(key:str).

Assume that this method '''def get_excel_global(key: str):''' is already defined in previous sections and usable now. Don't rewrite it. The code of this method is:
'''
from openpyxl import Workbook
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
    return cell.value
'''

The user first gives you the definition of variables.

You simply rewrite this piece into Python code. Each of your pieces is glued together in a Python file. This creates a complete program.
"""

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
This line begins with 'def' followed by the method name. The VBA code reads:
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
Method line begins with 'def' followed by the method name. The VBA code reads:
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

def prompt_dev_def(code: str, doc_block: str, var_code_py: str, sign_py, names: str) -> list:
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
        code=code, doc_block=doc_block, names_block=names_block, additional_instructions=additional_instructions
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

def request_dev(label: str, code: str, doc_block: str, var_code_py: str, sign_py, names: str) -> str:
    if label.startswith("++"):
        messages = prompt_dev_var(code, doc_block, names)
    else:
        messages = prompt_dev_def(code, doc_block, var_code_py, sign_py, names)
    response = get_response(messages, model=PROMPT_MODEL_CODE)
    return response

def request_sign(label: str, code: str, doc_block: str, var_code_py: str, names: str) -> str:
    messages = prompt_signatur(code, doc_block, var_code_py ,names)
    response = get_response(messages, model=PROMPT_MODEL_SIGN)
    return response

