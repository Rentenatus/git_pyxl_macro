

import math

import math

from math import isclose

from datetime import date

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

cache = None  # Initialize cache as None.  Type will be determined by usage.

rund_lx = 16
rund_tx = 16
rund_Dx = 16
rund_Cx = 16
rund_Nx = 16
rund_Mx = 16
rund_Rx = 16
max_Alter = 123

def initialize_cache():
    # Create a new dictionary object
    cache = {}

def act_qx(alter: int, sex: str, tafel: str, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial death probability (qx) for a given age, sex, and mortality table.

    Parameters:
    - alter (int): The age for which to calculate qx. Must be a non-negative integer.
    - sex (str): The sex of the individual ("M" for male, "F" for female). Case-insensitive.
    - tafel (str): The name of the mortality table to use. Currently supports "DAV1994_T" and "DAV2008_T".
    - geb_jahr (int, optional): The year of birth. Not currently used in the calculation.
    - rentenbeginnalter (int, optional): The age at which pension benefits begin. Not currently used.
    - schicht (int, optional): An integer representing the layer (Schicht) of the German pension system. Defaults to 1.

    Returns:
    - float: The actuarial death probability (qx). Returns 1.0 if the specified mortality table is not supported.
    """
    sex = sex.upper()
    if sex != "M":
        sex = "F"

    tafel_upper = tafel.upper()
    if tafel_upper in ["DAV1994_T", "DAV2008_T"]:
        s_tafelvektor = tafel_upper + "_" + sex

        # Get the named ranges
        v_tafeln_range = get_excel_global("v_Tafeln")
        m_tafeln_range = get_excel_global("m_Tafeln")

        # Find the index of the table in v_Tafeln
        try:
            index = v_tafeln_range.index(s_tafelvektor)
        except ValueError:
            return 1.0

        # Get the qx value from m_Tafeln
        try:
            qx_value = m_tafeln_range[alter][index]
            return float(qx_value) if qx_value is not None else 1.0
        except (IndexError, TypeError):
            return 1.0
    else:
        return 1.0

def v_lx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates and returns a vector representing the lx values (number of survivors at age x).

    Parameters:
    - endalter (int): The maximum age to calculate lx values up to. If -1, uses max_Alter.
    - sex (str): The sex of the population being modeled ("M" or "F").
    - tafel (str): Specifies the mortality table to use.
    - gebjahr (int, optional): The birth year of the cohort.
    - rentenbeginnalter (int, optional): The standard age at which pension payments begin.
    - schicht (int, optional): Represents the pension scheme layer (1-3).

    Returns:
    - list[float]: A list containing the lx values.
    """
    # Determine the calculation limit
    if endalter == -1:
        grenze = get_excel_global("max_Alter")
    else:
        grenze = endalter

    # Initialize the vector
    vek = [0.0] * (grenze + 1)  # Using 0.0 to ensure float type

    # Base case: lx(0) = 1,000,000
    vek[0] = 1000000.0

    # Iterative calculation
    for i in range(1, grenze + 1):
        qx = act_qx(i - 1, sex, tafel, gebjahr, rentenbeginnalter, schicht)
        vek[i] = vek[i - 1] * (1 - qx)

        # Round to the specified precision
        rund_lx = get_excel_global("rund_lx")
        vek[i] = round(vek[i], rund_lx)

    return vek

def act_lx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial life expectancy (lx) at a given age, considering sex, mortality table, and pension scheme layer.

    Parameters:
    - alter (int): The age for which to calculate the life expectancy (in years). This is the primary input and the index used to retrieve the value from the returned array.
    - sex (str): The sex of the individual ("m" for male, "w" for female). This influences the chosen mortality table.
    - tafel (str): The mortality table to use (e.g., "DAV2018", "GVG1994"). The availability of specific tables depends on the implementation of the underlying `v_lx` function.
    - gebjahr (int, optional): The year of birth. While not directly used in this function, it's passed to the `v_lx` function, suggesting it might be used for table adjustments or cohort effects within that function. If not provided, a default or current year is likely used internally within `v_lx`.
    - rentenbeginnalter (int, optional): The age at which pension payments begin. This is passed to the `v_lx` function and may influence the life expectancy calculation depending on the specifics of the `v_lx` implementation, potentially for annuitization calculations.
    - schicht (int, optional): Specifies the layer of the German pension system (1, 2, or 3). Defaults to 1. This parameter is critical as it ties the function directly to the broader German pension context.

    Returns:
    - float: The actuarial life expectancy (lx) at the given age, based on the provided parameters. The value represents the expected number of remaining years of life for a person of that age, sex, and pension scheme layer.

    Remarks:
    This function acts as a wrapper around the `v_lx` function. It retrieves an array of life expectancy values from `v_lx` and then returns the value corresponding to the input `alter`. Error handling (e.g., for invalid `sex` or `tafel` values) is likely handled within the `v_lx` function.
    """
    vek = v_lx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_tx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector of tx values (number of deaths at each age) based on lx values.

    Parameters:
    - endalter (int): Upper age limit for calculation. If -1, uses max_Alter constant.
    - sex (str): Gender identifier ("M" or "F").
    - tafel (str): Mortality table identifier.
    - gebjahr (int, optional): Year of birth.
    - rentenbeginnalter (int, optional): Age at which pension payments begin.
    - schicht (int, optional): Pension system layer (1, 2, or 3). Defaults to 1.

    Returns:
    - list: Array of tx values (number of deaths at each age), rounded to rund_tx precision.
    """
    # Get max_Alter from global names
    max_alter = get_excel_global('max_Alter')

    # Determine calculation limit
    grenze = max_alter if endalter == -1 else endalter

    # Initialize vector
    vek = [0] * grenze

    # Get lx values
    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    # Calculate tx values
    for i in range(grenze):
        vek[i] = v_temp_lx[i] - v_temp_lx[i + 1]
        # Get rund_tx from global names and round the value
        rund_tx = get_excel_global('rund_tx')
        vek[i] = round(vek[i], rund_tx)

    return vek

def act_tx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (tx) for a given age, sex, and mortality table.

    Parameters:
    - alter (int): The current age of the individual.
    - sex (str): The sex of the individual ("M" for male, "W" for female).
    - tafel (str): The identifier for the mortality table.
    - gebjahr (int, optional): The year of birth of the individual.
    - rentenbeginnalter (int, optional): The age at which pension payments begin.
    - schicht (int, optional): The "Schicht" (layer) of the German pension system. Defaults to 1.

    Returns:
    - float: The actuarial present value factor (tx) for the given age, sex, and mortality table.
    """
    vek = v_tx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_dx(endalter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector 'Dx' representing the present value of a life annuity due, considering various actuarial parameters.

    Parameters:
    - endalter (int): The final age to which the annuity is calculated. A value of -1 indicates calculation up to 'max_Alter' (a globally defined maximum age). This parameter defines the size of the resulting vector.
    - sex (str): A string indicating the sex of the annuitant. Used to select appropriate mortality rates. Example: "M" for male, "F" for female.
    - tafel (str): A string specifying the mortality table to use. Different tables represent different population demographics and mortality assumptions.
    - zins (float): The interest rate used for discounting future payments. This is crucial for calculating present values.
    - gebjahr (int, optional): The year of birth of the annuitant. May be used in conjunction with the mortality table.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin. Influences the calculation within the underlying mortality table function.
    - schicht (int, optional): An integer representing the "Schicht" (layer) of the German pension system (1, 2, or 3). This likely influences the selection of specific actuarial assumptions or rules. Defaults to 1.

    Returns:
    - list[float]: A list containing the calculated 'Dx' values. Each element represents the present value of a 1-Euro annuity payment due at that age, discounted at the specified interest rate and based on the provided mortality assumptions.
    """
    # Get global constants
    max_alter = get_excel_global('max_Alter')
    rund_dx = get_excel_global('rund_Dx')

    # Determine the boundary age
    grenze = max_alter if endalter == -1 else endalter

    # Initialize the vector
    vek = [0.0] * (grenze + 1)

    # Calculate the discount factor
    v = 1 / (1 + zins)

    # Get the life table values
    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    # Calculate Dx values
    for i in range(grenze + 1):
        vek[i] = v_temp_lx[i] * (v ** i)
        vek[i] = round(vek[i], rund_dx)

    return vek

cache = None

def act_dx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Dx' for a given age, sex, mortality table, interest rate, and pension scheme layer (Schicht).

    Parameters:
    - alter (int): The age for which to calculate the present value factor.
    - sex (str): The sex of the individual ("M" for male, "W" for female).
    - tafel (str): The mortality table to use (e.g., "DAV2018").
    - zins (float): The annual interest rate as a decimal (e.g., 0.05 for 5%).
    - gebjahr (int, optional): The year of birth. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which pension payments begin. Defaults to None.
    - schicht (int, optional): Indicates the layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The actuarial present value factor 'Dx'.
    """
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Dx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value is already in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store the result in the cache
        cache[s_key] = result
        return result

def v_cx(endalter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector of 'Cx' values, representing the present value of a life annuity-due.

    Parameters:
    - endalter (int): The final age to calculate Cx for. If -1, calculations are performed up to 'max_Alter' (a globally defined constant representing the maximum age).
    - sex (str): A string indicating the sex ("m" for male, "w" for female) for mortality table selection.
    - tafel (str): The name of the mortality table to use. This identifies the specific actuarial table used to determine mortality rates.
    - zins (float): The interest rate used for discounting future payments. Expressed as a decimal (e.g., 0.05 for 5%).
    - gebjahr (int, optional): The year of birth. Used in conjunction with 'Tafel' to select an appropriate mortality table. If not provided, a default table is assumed.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin. Used in conjunction with 'Tafel' and 'GebJahr'.
    - schicht (int, optional): An integer representing the layer/pillar of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - list[float]: A list containing the calculated 'Cx' values. Each element corresponds to an age. The values are rounded to the 'rund_Cx' decimal places (a globally defined constant).
    """
    # Get global constants
    max_alter = get_excel_global('max_Alter')
    rund_cx = get_excel_global('rund_Cx')

    # Determine the boundary
    grenze = max_alter if endalter == -1 else endalter

    # Initialize the vector
    vek = [0.0] * grenze

    # Calculate the discount factor
    v = 1 / (1 + zins)

    # Get the v_tx values
    v_temp_tx = v_tx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    # Calculate Cx values
    for i in range(grenze):
        vek[i] = v_temp_tx[i] * (v ** (i + 1))
        vek[i] = round(vek[i], rund_cx)

    return vek

def act_cx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the present value of a lifetime annuity (Actuarial Value of a Lifetime Annuity).

    Parameters:
    - alter (int): The age of the insured person.
    - sex (str): The gender of the insured person ("M" for male, "F" for female).
    - tafel (str): The mortality table used (e.g., "DAV2018", "RSA2018").
    - zins (float): The interest rate.
    - gebjahr (int, optional): The birth year of the insured person. Needed for some mortality tables.
    - rentenbeginnalter (int, optional): The age at which the pension begins.
    - schicht (int, optional): The layer of pension provision (default is 1). Represents the 3-layer model of German pension provision.
      1 = Statutory pension, 2 = Occupational pension/Riester, 3 = Private provision.

    Returns:
    - float: The present value of the lifetime annuity.
    """
    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Cx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value is already in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_cx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store the result in the cache
        cache[s_key] = result
        return result

def v_nx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the Nx values (present value of life annuity due) for actuarial calculations.

    Parameters:
    - sex (str): Gender of the insured ("M" for male, "F" for female)
    - tafel (str): Mortality table to use
    - zins (float): Discount interest rate
    - gebjahr (int, optional): Year of birth
    - rentenbeginnalter (int, optional): Age at which pension payments begin
    - schicht (int, optional): Layer of German pension system (default 1)

    Returns:
    - list[float]: Array of Nx values for each age up to max_Alter
    """
    max_alter = get_excel_global("max_Alter")
    rund_dx = get_excel_global("rund_Dx")

    vek = [0.0] * (max_alter + 1)  # Initialize array with max_alter + 1 elements

    v_temp_dx = v_dx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_alter] = v_temp_dx[max_alter]
    for i in range(max_alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_dx[i]
        vek[i] = round(vek[i], rund_dx)

    return vek

cache = None

def act_nx(alter: int, sex: str, tafel: str, zins: float, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial value 'Nx' representing the present value of a life annuity due,
    considering various actuarial and demographic parameters.

    Parameters:
    - alter (int): The age for which the actuarial value is calculated.
    - sex (str): The sex of the individual ('M' for male, 'F' for female).
    - tafel (str): The mortality table used for calculations.
    - zins (float): The interest rate used for present value calculations.
    - geb_jahr (int, optional): The year of birth. May influence the mortality table used.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin.
    - schicht (int, optional): Indicates the 'layer' or pillar of the German pension system being considered. Defaults to 1.

    Returns:
    - float: The actuarial value 'Nx' representing the present value of a life annuity due.
    """
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Nx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

    # Check if the value is already in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_nx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store the result in the cache
        cache[s_key] = result
        return result

def v_mx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates and returns the 'Mx' vector used in actuarial calculations for German pension provisions.

    Parameters:
    - sex (str): The sex of the annuitant ("M" for male, "F" for female)
    - tafel (str): The name/identifier of the mortality table to use
    - zins (float): The interest rate (as decimal, e.g., 0.05 for 5%)
    - gebjahr (int, optional): Year of birth of the annuitant
    - rentenbeginnalter (int, optional): Age at which annuity payments begin
    - schicht (int, optional): Layer of German pension system (1-3), defaults to 1

    Returns:
    - list[float]: The calculated 'Mx' values as a list
    """
    max_alter = get_excel_global('max_Alter')
    rund_mx = get_excel_global('rund_Mx')

    vek = [0.0] * (max_alter + 1)  # Create list with max_alter + 1 elements

    v_temp_cx = v_cx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_alter] = v_temp_cx[max_alter]
    for i in range(max_alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_cx[i]
        vek[i] = round(vek[i], rund_mx)

    return vek

cache = None

def act_mx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Mx' for a given age, sex, mortality table, interest rate, and pension scheme layer.
    This function utilizes caching to improve performance by storing previously calculated values.

    Parameters:
    - alter (int): The age for which to calculate the actuarial factor. Must be a valid age within the bounds of the mortality table.
    - sex (str): The sex of the individual. Expected values are "M" for male and "F" for female. Case sensitivity may be important.
    - tafel (str): The mortality table to use. This identifies the specific life table used for calculating probabilities of survival.
    - zins (float): The interest rate to use for discounting future payments. Expressed as a decimal (e.g., 0.05 for 5%).
    - gebjahr (int, optional): The year of birth. May be used in conjunction with the mortality table to determine the relevant life expectancy.
    - rentenbeginnalter (int, optional): The age at which the pension begins. Used for calculating the duration of payments.
    - schicht (int, optional): The layer (Schicht) of the German pension system. Default is 1. This parameter influences the calculation based on specific rules.

    Returns:
    - float: The actuarial present value factor 'Mx' used to calculate the present value of a future pension payment.

    Remarks:
    This function is designed for calculations within the context of German pension actuarial work.
    It leverages a caching mechanism to store previously calculated 'Mx' values, significantly improving performance for repeated calculations.
    The 'v_mx' function is a core dependency for the actual calculation.
    The key used for caching is constructed from the input parameters to ensure uniqueness.
    """
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Mx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value is already in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store result in cache
        cache[s_key] = result
        return result

def v_rx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> list:
    """
    Calculates the vector of 'Rx' values, representing the present value of a life annuity due for each age.

    Parameters:
    - sex (str): The sex of the annuitant ("m" for male, "w" for female). This influences mortality rates and therefore annuity calculations.
    - tafel (str): The mortality table to use for calculating life expectancy. Different tables represent different population demographics and impact calculations.
    - zins (float): The interest rate used for discounting future payments. Higher interest rates reduce the present value of the annuity.
    - gebjahr (int, optional): The year of birth of the annuitant. Used in conjunction with 'tafel' to determine relevant mortality rates for specific ages.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin. This impacts the number of payments and their present value.
    - schicht (int, optional): Indicates the layer (Schicht) of the German pension system the calculation applies to. Defaults to 1. This parameter likely influences the specific annuity assumptions or factors used.

    Returns:
    - list: A list containing the calculated 'Rx' values for each age from 0 to 'max_Alter'. Each element represents the present value of an annuity-due at that age, rounded to 'rund_Rx' decimal places.
    """
    max_alter = get_excel_global("max_Alter")
    rund_rx = get_excel_global("rund_Rx")

    vek = [0.0] * (max_alter + 1)
    v_temp_mx = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_alter] = v_temp_mx[max_alter]
    for i in range(max_alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_mx[i]
        vek[i] = round(vek[i], rund_rx)

    return vek

cache = None

def act_rx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (Rx) for a given age, sex, mortality table, interest rate, and pension scheme layer.

    Parameters:
    - alter (int): The age for which to calculate the present value factor.
    - sex (str): The sex of the annuitant ("M" for male, "F" for female).
    - tafel (str): The mortality table to use (e.g., "DAV2018").
    - zins (float): The interest rate used for the present value calculation.
    - gebjahr (int, optional): The year of birth. Used in conjunction with the mortality table to determine life expectancy.
    - rentenbeginnalter (int, optional): The age at which the pension begins. Influences the calculation within the underlying v_Rx function.
    - schicht (int, optional): The layer of the German pension system (1, 2, or 3). Defaults to 1. This parameter influences the overall pension calculation scheme.

    Returns:
    - float: The actuarial present value factor (Rx).

    Remarks:
    This function calculates the actuarial present value factor (Rx) which is a crucial component in calculating pension benefits.
    It utilizes a caching mechanism to improve performance by storing previously calculated values.

    The function retrieves the value from a dictionary cache if it exists. If not, it calculates the factor using the v_Rx function
    and then stores the result in the cache for future use.

    The parameter 'schicht' directly relates to the German three-pillar pension system ("3-Schichten-Modell der Altersvorsorge").
    The function accounts for differing calculation approaches according to the layer chosen (1 = state pension, 2 = supplementary company/Riester, 3 = private).
    """
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Rx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value already exists in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_rx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store the result in the cache
        cache[s_key] = result
        return result

def act_altersberechnung(gebdat: date, berdat: date, methode: str) -> int:
    """
    Calculates a person's age based on their date of birth and a reference date.
    Supports two calculation methods: calendar year method ("K") and half-year method ("H").

    Parameters:
    - gebdat (date): The date of birth of the person.
    - berdat (date): The reference date for the age calculation.
    - methode (str): The age calculation method ("K" for calendar year, "H" for half-year).

    Returns:
    - int: The calculated age in years.
    """
    if methode != "K":
        methode = "H"

    j_gd = gebdat.year
    j_bd = berdat.year
    m_gd = gebdat.month
    m_bd = berdat.month

    if methode == "K":
        return j_bd - j_gd
    else:  # Default to "H" method
        return int(j_bd - j_gd + 1 / 12 * (m_bd - m_gd + 5))

def create_cache_key(art: str, alter: int, sex: str, tafel: str, zins: float, geb_jahr: int, rentenbeginnalter: int, schicht: int) -> str:
    """
    Creates a unique key for caching pension calculation results. This key is composed of the input parameters,
    allowing the macro to efficiently store and retrieve previously calculated values, avoiding redundant computations.

    Parameters:
    - art (str): String representing the type of pension calculation (e.g., "Riester", "gesetzlich", "betrieblich").
    - alter (int): Integer representing the current age of the individual.
    - sex (str): String representing the gender of the individual ("m" for male, "w" for female, or other relevant codes).
    - tafel (str): String identifying the mortality table (Lebensdauer-Tafel) used in the calculation. Different tables represent different demographic assumptions.
    - zins (float): Double representing the interest rate used for discounting future cash flows.
    - geb_jahr (int): Integer representing the year of birth of the individual.
    - rentenbeginnalter (int): Integer representing the age at which the pension payments begin.
    - schicht (int): Integer representing the pension "layer" (Schicht) according to the German three-layer pension system (1, 2, or 3). This categorizes the type of pension provision.

    Returns:
    - str: A string that uniquely identifies the combination of input parameters. This key is used for caching.

    Remarks:
    This function is crucial for optimizing performance within the larger pension calculation macro.
    By creating a unique key based on all relevant input parameters, the macro can store and retrieve results from a cache,
    avoiding the need to recalculate the same pension value repeatedly. The underscore "_" is used as a delimiter
    to ensure a valid and easily parsable key. Understanding the German pension system's 3-Schichten-Modell is helpful
    to interpreting the meaning of the 'schicht' parameter.
    """
    return f"{art}_{alter}_{sex}_{tafel}_{zins}_{geb_jahr}_{rentenbeginnalter}_{schicht}"

def act_ax_k(alter: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial function 'ax_k', representing the present value of a whole life insurance benefit payable at the beginning of the year of death, adjusted by a factor 'k'.

    Parameters:
    - alter (int): The current age of the individual.
    - sex (str): The sex of the individual ("M" for male, "F" for female).
    - tafel (str): The mortality table to be used.
    - zins (float): The interest rate (expressed as a decimal, e.g., 0.05 for 5%).
    - k (int): An adjustment factor applied to the benefit. If k <= 0, the function returns 0.
    - gebjahr (int, optional): The year of birth. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which benefits begin. Defaults to None.
    - schicht (int, optional): Represents the layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The calculated value of 'ax_k'. Returns 0 if k <= 0.
    """
    if k > 0:
        return act_nx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_abzugsglied(k, zins)
    else:
        return 0.0

def act_axn_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    if k > 0:
        return (act_nx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_nx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_abzugsglied(k, zins) * (1 - act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht))
    else:
        return 0.0

def act_nax_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    if k > 0:
        return act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) * act_ax_k(alter + n, sex, tafel, zins, k, gebjahr, rentenbeginnalter, schicht)
    else:
        return 0.0

def act_ngr_ax(alter: int, n: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the average present value of a life annuity-due with a period of 'n' years.
    Specifically, it computes the present value of a series of payments made for 'n' years,
    assuming payments begin immediately (annuity-due).

    Parameters:
    - alter (int): The current age of the individual.
    - n (int): The number of years for which the annuity is paid.
    - sex (str): The gender of the individual ("M" for male, "F" for female).
    - tafel (str): The mortality table used for calculations.
    - zins (float): The interest rate used for present value calculations.
    - gebjahr (int, optional): The year of birth of the individual.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin.
    - schicht (int, optional): A layer or level within the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The calculated average present value of the annuity-due.
    """
    return (act_mx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_mx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

def act_ngr_ex(alter: int, n: int, sex: str, tafel: str, zins: float, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the growth rate of the death probability (qx) over a period of 'n' years.

    Parameters:
    - alter (int): The current age of the individual.
    - n (int): The number of years into the future to calculate the growth rate.
    - sex (str): The sex of the individual ("M" for male, "F" for female).
    - tafel (str): The mortality table being used.
    - zins (float): The interest rate used in the calculation.
    - geb_jahr (int, optional): The year of birth. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which a pension or annuity begins. Defaults to None.
    - schicht (int, optional): Indicates the layer of the German pension system being considered. Defaults to 1.

    Returns:
    - float: The growth rate of the death probability (qx) over 'n' years.
    """
    return act_dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

def act_ag_k(g: int, zins: float, k: int) -> float:
    v = 1 / (1 + zins)
    if k > 0:
        if zins > 0:
            return (1 - v ** g) / (1 - v) - act_abzugsglied(k, zins) * (1 - v ** g)
        else:
            return g
    else:
        return 0

def act_abzugsglied(k: int, zins: float) -> float:
    """
    Calculates the 'Abzugsglied' (reduction factor) used in actuarial calculations, particularly in life insurance and pension mathematics.

    Parameters:
    - k (int): An integer representing the number of years or periods for which the reduction applies.
    - zins (float): A float representing the interest rate.

    Returns:
    - float: The calculated 'Abzugsglied' (reduction factor). Returns 0 if k is 0.
    """
    result = 0.0
    if k > 0:
        for l in range(0, k):
            result += l / k / (1 + l / k * zins)
        result = result * (1 + zins) / k
    return result