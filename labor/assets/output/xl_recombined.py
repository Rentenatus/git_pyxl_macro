

import numpy as np

import numpy as np


from openpyxl import Workbook
from openpyxl.utils import range_boundaries

xl_workbook: Workbook # The Excel workbook object
xl_names: dict [str, str] # The dictionary of named ranges

def get_excel_global(key: str):
    """
    Retrieves the value from the Excel workbook based on a named range key.
    
    Parameters:
    - key (str): The name of the range to look up.
    
    Returns:
    - The value of the cell referenced by the named range, or None if not found.
    """
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

# Variable to hold a caching object.  The specific type of object
# assigned to 'cache' is determined by the surrounding code and may
# include a Dictionary, Collection, or cached Excel range data.
cache = None

rund_lx = 16
rund_tx = 16
rund_Dx = 16
rund_Cx = 16
rund_Nx = 16
rund_Mx = 16
rund_Rx = 16
max_Alter = 123

def initialize_cache():
    """
    Creates a new dictionary object and assigns it to the 'cache' variable.
    """
    cache = {}  # Python's equivalent of Scripting.Dictionary
    return cache

def Act_qx(Alter: int, Sex: str, Tafel: str, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial mortality probability (qx) for a given age, sex, and mortality table.

    Parameters:
    - Alter (int): The age for which to calculate qx.
    - Sex (str): The sex of the individual ("M" for male, "F" for female). Case insensitive.
    - Tafel (str): The mortality table identifier. Currently supports "DAV1994_T" and "DAV2008_T". Case insensitive.
    - GebJahr (int, optional): Year of birth. Not currently used. Defaults to None.
    - Rentenbeginnalter (int, optional): Age at which pension benefits begin. Not currently used. Defaults to None.
    - Schicht (int, optional): Layer of the German pension system. Not currently used. Defaults to 1.

    Returns:
    - The actuarial mortality probability (qx) as a float. Returns 1 if the specified mortality table is not supported.

    Raises:
    - KeyError: If the worksheet or ranges are not found in the Excel workbook.
    """
    try:
        # Assuming "Tafeln" is the sheet name
        ws = xl_workbook["Tafeln"]
    except KeyError:
        raise KeyError("Worksheet 'Tafeln' not found in xl_workbook.")

    if Sex.upper() != "M":
        Sex = "F"

    try:
        match_range = ws["v_Tafeln"]
        qx_range = ws["m_Tafeln"]
    except KeyError:
        raise KeyError("Ranges 'v_Tafeln' or 'm_Tafeln' not found in worksheet 'Tafeln'.")
    
    table_sex_string = Tafel.upper() + "_" + Sex
    
    # Find the match in v_Tafeln using a loop for compatibility with Excel's MATCH
    match_index = None
    for i in range(1, len(match_range) + 1):
        if str(match_range[i]).upper() == table_sex_string:
            match_index = i
            break

    if match_index is None:
      return 1.0
    
    qx_value = qx_range[match_index][Alter]
    
    return qx_value

def v_lx(Endalter: int, Sex: str, Tafel: str, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> list[float]:
    """
    Calculates the vector 'lx', representing the number of survivors at each age.

    Parameters:
    - Endalter (int): The maximum age to calculate the survivor vector up to. If -1, the vector is calculated up to the maximum defined age ('max_Alter').
    - Sex (str): The sex of the individual ("M" for male, "F" for female).
    - Tafel (str): The mortality table to use (e.g., "DAV2018", "Hensel").  This determines the base mortality rates.
    - GebJahr (int, optional): The year of birth.  May be used within the 'Act_qx' function to adjust mortality rates based on cohort effects.
    - Rentenbeginnalter (int, optional): The age at which pension payments begin. May be used within 'Act_qx'.
    - Schicht (int, optional): An integer representing the layer (Schicht) of the German pension system (1, 2, or 3).  Defaults to 1. Used within the 'Act_qx' function, likely influencing mortality assumptions based on the pillar of retirement provision.

    Returns:
    - A list representing the 'lx' vector.  Each element of the array corresponds to the number of survivors at that age.  The first element (index 0) represents age 0.
    """
    try:
        max_Alter = get_excel_global("max_Alter")
        rund_lx = get_excel_global("rund_lx")
    except KeyError as e:
        raise KeyError(f"Required global variable not found: {e}")
    
    if Endalter == -1:
        Grenze = max_Alter
    else:
        Grenze = Endalter

    vek = [0.0] * (Grenze + 1)  # Initialize the list with the correct size
    vek[0] = 1000000.0

    for i in range(1, Grenze + 1):
        vek[i] = vek[i - 1] * (1 - Act_qx(i - 1, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht))
        vek[i] = round(vek[i], rund_lx)

    return vek

def Act_lx(Alter: int, Sex: str, Tafel: str, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial lifetime expectation (`lx`) at a given age.

    Parameters:
        Alter (int): The current age.
        Sex (str): The sex of the individual ("M" or "W").
        Tafel (str): The name of the mortality table.
        GebJahr (int, optional): The birth year. Defaults to None.
        Rentenbeginnalter (int, optional): The retirement age. Defaults to None.
        Schicht (int, optional): The pension scheme layer. Defaults to 1.

    Returns:
        float: The actuarial lifetime expectation (`lx`).
    """
    vek = v_lx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    return vek[Alter]

def v_tx(Endalter: int, Sex: str, Tafel: str, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the vector of tx values (number of deaths) for a given cohort.

    Parameters:
        Endalter (int): The upper age limit for the calculation.
        Sex (str): The sex of the cohort ("m" or "w").
        Tafel (str): The mortality table to be used.
        GebJahr (Optional[int]): The birth year of the cohort.
        Rentenbeginnalter (Optional[int]): The age at which pension payments begin.
        Schicht (Optional[int]): The pension scheme layer (default is 1).

    Returns:
        np.ndarray: A NumPy array containing the calculated tx values.
    """

    if Endalter == -1:
        Grenze = get_excel_global("max_Alter")
    else:
        Grenze = Endalter

    vek = np.zeros(Grenze)
    v_Temp_lx = v_lx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)

    for i in range(Grenze - 1):
        vek[i] = v_Temp_lx[i] - v_Temp_lx[i + 1]
        vek[i] = round(vek[i], get_excel_global("rund_tx"))

    return vek

def Act_tx(Alter: int, Sex: str, Tafel: str, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (tx) for a given age, sex, and life table,
    considering potential adjustments for birth year, retirement age, and pension scheme layer (Schicht).
    """
    vek = v_tx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    Act_tx = vek[Alter]
    return Act_tx

def v_Dx(Endalter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the vector of 'Dx' values, representing the present values of surviving to each age.

    Parameters:
        Endalter (int): The maximum age to calculate the Dx values for. If -1, calculations are performed up to the defined maximum age (<see cref="max_Alter"/>).
        Sex (str): The sex of the individual ("M" for male, "F" for female). This influences the mortality table used.
        Tafel (str): The mortality table to use. Identifies the specific mortality assumptions.
        Zins (float): The discount rate (interest rate) used to calculate present values.
        GebJahr (int, optional): The year of birth. May be used by the underlying 'v_lx' function to determine the appropriate mortality table.
        Rentenbeginnalter (int, optional): The age at which a pension or annuity begins. May be used by the underlying 'v_lx' function.
        Schicht (int, optional): An integer indicating the pension scheme layer (1, 2, or 3) within the German pension system (Altervorsorge). Defaults to 1.

    Returns:
        np.ndarray: A NumPy array containing the calculated 'Dx' values for each age from 0 up to 'Endalter' (or 'max_Alter' if 'Endalter' is -1). Each element of the array represents the present value of surviving to that age.
    
    """
    
    if Endalter == -1:
        Grenze = get_excel_global("max_Alter")
    else:
        Grenze = Endalter
        
    vek = np.zeros(Grenze + 1)
    v = 1 / (1 + Zins)
    
    v_Temp_lx = v_lx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    
    for i in range(Grenze + 1):
        vek[i] = v_Temp_lx[i] * (v ** i)
        vek[i] = round(vek[i], get_excel_global("rund_Dx"))
        
    return vek

cache = {}  # Initialize the cache dictionary

def Act_Dx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Dx' for a given age, sex, mortality table, interest rate, and pension scheme layer.
    """
    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    sKey = CreateCacheKey("Dx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    # Check if the value already exists in the cache
    if sKey in cache:
        return cache[sKey]
    else:
        vek = v_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
        result = vek[Alter]

        # Store the result in the cache
        cache[sKey] = result

        return result

def v_Cx(Endalter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the vector of 'Cx' values, representing the present value of a life annuity-immediate.

    Parameters:
        Endalter (int): The final age to calculate the Cx values for.  A value of -1 indicates calculation up to the maximum supported age (max_Alter).
        Sex (str): The sex of the annuitant ("M" for male, "F" for female). Affects mortality table selection.
        Tafel (str): The mortality table to use for calculations.  Different tables represent different population characteristics and mortality rates.
        Zins (float): The annual interest rate used for discounting future payments. Expressed as a decimal (e.g., 0.05 for 5%).
        GebJahr (int, optional): The year of birth. May be used by some mortality tables.
        Rentenbeginnalter (int, optional): The age at which the annuity payments begin.
        Schicht (int, optional): An integer representing the layer (Schicht) of the German pension system (1, 2, or 3). Defaults to 1. This parameter influences the calculation of the 'tx' values used within the function, accounting for different regulations and tax treatments for each layer.

    Returns:
        list[float]: A list containing the calculated 'Cx' values.  Each element of the list corresponds to a specific age, representing the present value of an annuity payment at that age.
    """
    # erzeugt Vektor der Cx
    vek = []
    i = 0
    
    Grenze = 0
    
    if Endalter == -1:
        Grenze = get_excel_global('max_Alter')
    else:
        Grenze = Endalter
    
    # ReDim vek(Grenze)
    v = 1 / (1 + Zins)
    
    v_Temp_tx = v_tx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
            
    for i in range(Grenze):
        vek.append(v_Temp_tx[i] * v ** (i + 1))
        vek[i] = round(vek[i], get_excel_global('rund_Cx'))
    
    return vek

cache = {}  # Initialize the cache dictionary

def Act_Cx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Cx' for a given age, sex, mortality table, interest rate, and pension scheme layer.
    """
    # Check if the cache dictionary is initialized
    if cache is None:
        initialize_cache()

    sKey = CreateCacheKey("Cx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    # Check if the value is already in the cache
    if sKey in cache:
        return cache[sKey]
    else:
        vek = v_Cx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
        result = vek[Alter]

        # Store the result in the cache
        cache[sKey] = result
        return result

def v_Nx(Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the 'Nx' vector, representing the present value of a life annuity due, contingent on survival to age x.

    Parameters:
        Sex (str): The sex of the annuitant ("M" for male, "F" for female).
        Tafel (str): The mortality table being used (e.g., "DAV2018", "TVA2018"). Determines the survival probabilities.
        Zins (float): The interest rate (as a decimal, e.g., 0.05 for 5%).
        GebJahr (int, optional): The year of birth of the annuitant. Used in conjunction with the mortality table to determine survival probabilities. Defaults to None.
        Rentenbeginnalter (int, optional): The age at which the annuity payments begin. Impacts the number of periods considered in the calculation. Defaults to None.
        Schicht (int, optional): Indicates the layer (Schicht) of the German pension system (1, 2, or 3). This likely impacts the tax treatment or specific rules applied to the calculation. Defaults to 1.

    Returns:
        list: A list representing the 'Nx' vector. Each element of the list corresponds to the present value of a life annuity due contingent on survival to a given age. The list is indexed by age.
    """
    vek = [0.0] * max_Alter
    v_Temp_Dx = v_Dx(-1, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    vek[max_Alter - 1] = v_Temp_Dx[max_Alter - 1]

    for i in range(max_Alter - 2, -1, -1):
        vek[i] = vek[i + 1] + v_Temp_Dx[i]
        vek[i] = round(vek[i], rund_Dx)

    return vek

cache = {}  # Initialize the cache here

def Act_Nx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the probability that a person of age 'Alter' is still alive (Actuarial Function Nx).
    """
    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    sKey = CreateCacheKey("Nx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    # Check if the value already exists in the cache
    if sKey in cache:
        return cache[sKey]
    else:
        vek = v_Nx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
        result = vek[Alter]

        # Store the result in the cache
        cache[sKey] = result
        return result

def v_Mx(Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the 'Mx' vector, representing the present value of a life annuity due, considering various actuarial parameters.

    Parameters:
        Sex (str): The sex of the individual ('M' for male, 'W' for female). This impacts mortality assumptions.
        Tafel (str): The mortality table to use for calculations (e.g., 'DAV2018', 'TGV2018'). Specifies the expected lifespan based on sex.
        Zins (float): The interest rate used for present value calculations. A key factor in determining the value of future payments.
        GebJahr (int, optional): The year of birth of the individual. May be used in conjunction with the 'Tafel' to refine mortality assumptions. Defaults to None.
        Rentenbeginnalter (int, optional): The age at which the annuity payments begin. Impacts the number of payments to be discounted. Defaults to None.
        Schicht (int, optional): Indicates the layer of the German pension system (1, 2, or 3). This parameter influences the calculation context, potentially impacting tax or benefit considerations. Defaults to 1.

    Returns:
        list: A list ('vek') containing the calculated 'Mx' values for each age up to 'max_Alter'. Each element represents the present value of remaining life years from that age onwards. Values are rounded to 'rund_Mx' decimal places.
    """
    vek = [0.0] * (max_Alter + 1)
    
    v_Temp_Cx = v_Cx(-1, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    
    vek[max_Alter] = v_Temp_Cx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_Temp_Cx[i]
        vek[i] = round(vek[i], rund_Mx)
    
    return vek

cache = {}  # Initialize the cache dictionary

def Act_Mx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the present value of the life insurance function Mx for a given age.
    """
    # Check if the cache is initialized
    if not cache:
        initialize_cache()

    sKey = CreateCacheKey("Mx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    # Check if the value is already in the cache
    if sKey in cache:
        return cache[sKey]
    else:
        vek = v_Mx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
        result = vek[Alter]

        # Store the result in the cache
        cache[sKey] = result
        return result

def v_Rx(Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1):
    """
    Calculates the vector 'Rx', representing the present value of a life annuity-due, crucial for pension benefit calculations.

    Parameters:
        Sex (str): Gender of the annuitant ("M" for male, "W" for female). Influences mortality rates used in the underlying calculation.
        Tafel (str): Mortality table identifier. Specifies the mortality table used to determine probabilities of survival. Different tables reflect different population characteristics and assumptions.
        Zins (float): Interest rate (as a decimal). Used to discount future cash flows to their present value.
        GebJahr (int, optional): Year of birth of the annuitant. May be used in conjunction with the 'Tafel' to select the appropriate mortality rates. Defaults to None.
        Rentenbeginnalter (int, optional): Age at which the annuity payments begin. Determines the length of the annuity period. Defaults to None.
        Schicht (int, optional): The layer of the German pension system (1, 2, or 3). Defaults to 1. This parameter may influence the specific calculation logic or assumptions used, reflecting the different rules governing each layer (gesetzliche Rente, betriebliche Altersvorsorge, private Vorsorge).

    Returns:
        list[float]: A list representing the 'Rx' vector. Each element of the vector corresponds to the present value of a life annuity-due paid at a specific age, starting from 'Rentenbeginnalter' and continuing until the maximum age ('max_Alter'). The values are rounded to a specified decimal place ('rund_Rx').
    """
    # erzeugt Vektor der Rx
    vek = [0.0] * max_Alter

    v_Temp_Mx = v_Mx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    vek[max_Alter - 1] = v_Temp_Mx[max_Alter - 1]
    for i in range(max_Alter - 2, -1, -1):
        vek[i] = vek[i + 1] + v_Temp_Mx[i]
        vek[i] = round(vek[i], rund_Rx)

    return vek

cache = {}  # Initialize the cache dictionary

def Act_Rx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Rx' for a given age, sex, mortality table, interest rate, and pension scheme layer.
    """
    # Check if the cache has been initialized
    if cache is None:
        initialize_cache()

    sKey = CreateCacheKey("Rx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

    # Check if the value is already in the cache
    if sKey in cache:
        return cache[sKey]
    else:
        vek = v_Rx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
        result = vek[Alter]

        # Store the result in the cache
        cache[sKey] = result
        return result

def Act_Altersberechnung(GebDat, BerDat, Methode):
    """
    Calculates age based on a given date of birth and a reference date, 
    using either the calendar year method or the half-year method.

    Parameters:
        GebDat (date): The date of birth.
        BerDat (date): The reference date.
        Methode (str): The calculation method ("K" for calendar year, otherwise half-year).

    Returns:
        int: The calculated age in years.
    """
    if Methode != "K":
        Methode = "H"

    J_GD = GebDat.year
    J_BD = BerDat.year
    M_GD = GebDat.month
    M_BD = BerDat.month

    if Methode == "K":
        return J_BD - J_GD
    else:  # Methode == "H"
        return int(J_BD - J_GD + 1 / 12 * (M_BD - M_GD + 5))

def create_cache_key(art: str, alter: int, sex: str, tafel: str, zins: float, geb_jahr: int, rentenbeginnalter: int, schicht: int) -> str:
    """
    Creates a unique key for caching pension calculation results. This key is constructed from the input parameters, allowing the macro to efficiently retrieve previously calculated values instead of recomputing them.

    Parameters:
    - art (str): String representing the type of pension calculation. (e.g., "Rentenberechnung", "Kapitalberechnung").
    - alter (int): Integer representing the current age of the person.
    - sex (str): String representing the sex of the person ("m" for male, "w" for female).
    - tafel (str): String representing the mortality table used for the calculation. Different tables represent varying life expectancy assumptions.
    - zins (float): Double representing the interest rate used in the calculation.
    - geb_jahr (int): Integer representing the year of birth.
    - rentenbeginnalter (int): Integer representing the age at which the pension begins.
    - schicht (int): Integer representing the pension pillar/layer (Schicht) being calculated. This relates to the German three-pillar pension system (1st, 2nd, or 3rd layer).

    Returns:
    - A string representing the unique cache key.

    Remarks:
    This function is crucial for performance optimization. By generating a unique key based on the input parameters, the macro can store and retrieve results from a cache (e.g., a Dictionary object) instead of recalculating them repeatedly. The underscore "_" is used as a delimiter to create a readable and easily parsable key.
    """
    return f"{art}_{alter}_{sex}_{tafel}_{zins}_{geb_jahr}_{rentenbeginnalter}_{schicht}"

def Act_ax_k(Alter: int, Sex: str, Tafel: str, Zins: float, k: int, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the factor 'ax_k' for converting pension claims.
    """
    if k > 0:
        Act_ax_k = Act_Nx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) - Act_Abzugsglied(k, Zins)
    else:
        Act_ax_k = 0
    return Act_ax_k

def Act_axn_k(Alter: int, n: int, Sex: str, Tafel: str, Zins: float, k: int, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the factor axn,k, which indicates the probability that a person
    will still be alive at age 'Alter' and die 'n' years later, reduced by a
    deduction. This factor is used in the calculation of annuities and life
    insurance.

    Parameters:
    - Alter (int): The current age of the person in years.
    - n (int): The number of years into the future to consider.
    - Sex (str): The gender of the person ("m" for male, "w" for female).
    - Tafel (str): The mortality table used (e.g., DAV, HAB).
    - Zins (float): The interest rate as a decimal (e.g., 0.05 for 5%).
    - k (int): A parameter that indicates the number of periods for a specific
      annuity or insurance. If k <= 0, the return value is 0.
    - GebJahr (int, optional): The year of birth of the person. May be used
      in the called functions.
    - Rentenbeginnalter (int, optional): The age at which the annuity is to
      begin. May be used in the called functions.
    - Schicht (int, optional): The layer of pension provision (1, 2 or 3).
      This serves to distinguish between different pension areas. Default
      value is 1.

    Returns:
    - The calculated factor axn,k as a float. Returns 0 if k is less than or
      equal to 0.
    """
    if k > 0:
        return (Act_Nx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) - Act_Nx(Alter + n, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) - Act_Abzugsglied(k, Zins) * (1 - Act_Dx(Alter + n, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht))
    else:
        return 0

def Act_nax_k(Alter: int, n: int, Sex: str, Tafel: str, Zins: float, k: int, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the present value of a deferred annuity that begins in 'n' years,
    taking into account gender, age, mortality table, interest rate, and a deduction factor 'k'.
    """
    if k > 0:
        Act_nax_k = Act_Dx(Alter + n, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) * act_ax_k(Alter + n, Sex, Tafel, Zins, k, GebJahr, Rentenbeginnalter, Schicht)
    else:
        Act_nax_k = 0
    return Act_nax_k

def Act_nGrAx(Alter: int, n: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the average present value of a life annuity-due of amount 1,
    payable for 'n' years, starting at age 'Alter' + 'n'.
    """
    return (Act_Mx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) - Act_Mx(Alter + n, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

def Act_nGrEx(Alter: int, n: int, Sex: str, Tafel: str, Zins: float, GebJahr: int = None, Rentenbeginnalter: int = None, Schicht: int = 1) -> float:
    """
    Calculates the growth rate of the death probability (q<sub>x</sub>) over 'n' years.
    This function determines how the likelihood of death changes between ages 'Alter' and 'Alter + n'.
    It is a core component in actuarial calculations related to life insurance and pension benefit estimations.
    """
    return Act_Dx(Alter + n, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)

def Act_ag_k(g: int, Zins: float, k: int) -> float:
    """
    Calculates a factor used in actuarial calculations.

    Parameters:
    g (int): The number of periods.
    Zins (float): The interest rate per period.
    k (int): A parameter controlling the size of a deduction.

    Returns:
    float: The calculated factor.
    """
    v = 1 / (1 + Zins)
    if k > 0:
        if Zins > 0:
            Act_ag_k = (1 - v ** g) / (1 - v) - Act_Abzugsglied(k, Zins) * (1 - v ** g)
        else:
            Act_ag_k = g
    else:
        Act_ag_k = 0
    return Act_ag_k

def Act_Abzugsglied(k: int, Zins: float) -> float:
    """
    Calculates the 'Abzugsglied' (reduction factor) used in actuarial calculations.

    Parameters:
    k (int): The number of periods or years over which the reduction applies.
    Zins (float): The interest rate.

    Returns:
    float: The calculated 'Abzugsglied'. Returns 0 if k is not positive.
    """
    Act_Abzugsglied = 0
    if k > 0:
        for l in range(k):
            Act_Abzugsglied = Act_Abzugsglied + l / k / (1 + l / k * Zins)
        Act_Abzugsglied = Act_Abzugsglied * (1 + Zins) / k
    return Act_Abzugsglied