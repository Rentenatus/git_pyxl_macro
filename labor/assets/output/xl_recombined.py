

import numpy as np

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
    return cache

def act_qx(alter: int, sex: str, tafel: str, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial death probability (qx) for a given age, sex, and mortality table.
    """
    sex = sex.upper()
    if sex != "M":
        sex = "F"

    if tafel.upper() in ("DAV1994_T", "DAV2008_T"):
        s_tafelvektor = tafel.upper() + "_" + sex
        try:
            # Assuming "v_Tafeln" is a horizontal range
            match_index = 0  # Excel Match returns 1-based index
            v_tafeln_values = get_excel_global("v_Tafeln")
            
            if v_tafeln_values is None:
                raise ValueError("v_Tafeln not found in Excel.")

            try:
                match_index = v_tafeln_values.index(s_tafelvektor) + 1 # Python lists are 0-indexed, Excel starts at 1.
            except ValueError:
                return 1.0 # Raise error 1

            # Assuming "m_Tafeln" is a vertical range
            m_tafeln_values = get_excel_global("m_Tafeln")
            if m_tafeln_values is None:
                raise ValueError("m_Tafeln not found in Excel.")
            
            qx = m_tafeln_values[alter]
            return qx
            
        except (KeyError, ValueError) as e:
            return 1.0 # error 1
    else:
        return 1.0 # error 1

def v_lx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates and returns a vector representing the lx values (number of survivors at age x).

    Parameters:
        endalter (int): The maximum age to calculate lx values up to.
                         If endalter is -1, the calculation will extend up to max_Alter.
        sex (str): The sex of the population being modeled ("M" or "F").
        tafel (str): Specifies the mortality table to use.
        gebjahr (int, optional): The birth year of the cohort. Defaults to None.
        rentenbeginnalter (int, optional): The standard age at which pension payments begin. Defaults to None.
        schicht (int, optional): Represents the pension scheme layer. Defaults to 1.

    Returns:
        list: A list containing the lx values.
    """
    if endalter == -1:
        grenze = get_excel_global("max_Alter")
    else:
        grenze = endalter

    vek = [0] * grenze
    vek[0] = 1000000

    rund_lx = get_excel_global("rund_lx")

    for i in range(1, grenze):
        vek[i] = vek[i - 1] * (1 - act_qx(i - 1, sex, tafel, gebjahr, rentenbeginnalter, schicht))
        vek[i] = round(vek[i], rund_lx)

    return vek

def act_lx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial life expectancy (lx) at a given age, considering sex, mortality table, and pension scheme layer.

    Parameters:
        alter (int): The age for which to calculate the life expectancy (in years).  This is the primary input and the index used to retrieve the value from the returned array.
        sex (str): The sex of the individual ("m" for male, "w" for female).  This influences the chosen mortality table.
        tafel (str): The mortality table to use (e.g., "DAV2018", "GVG1994"). The availability of specific tables depends on the implementation of the underlying `v_lx` function.
        gebjahr (int, optional): The year of birth. While not directly used in this function, it's passed to the `v_lx` function, suggesting it might be used for table adjustments or cohort effects within that function.  If not provided, a default or current year is likely used internally within `v_lx`.
        rentenbeginnalter (int, optional): The age at which pension payments begin. This is passed to the `v_lx` function and may influence the life expectancy calculation depending on the specifics of the `v_lx` implementation, potentially for annuitization calculations.
        schicht (int, optional):  Specifies the layer of the German pension system (1, 2, or 3).  Defaults to 1.  This parameter is critical as it ties the function directly to the broader German pension context.  See the "Relationship to German Pension System" section below for more detail.

    Returns:
        float: The actuarial life expectancy (lx) at the given age, based on the provided parameters.  The value represents the expected number of remaining years of life for a person of that age, sex, and pension scheme layer.

    Remarks:
        This function acts as a wrapper around the `v_lx` function. It retrieves an array of life expectancy values from `v_lx` and then returns the value corresponding to the input `Alter`.  Error handling (e.g., for invalid `Sex` or `Tafel` values) is likely handled within the `v_lx` function.

    See Also:
        v_lx

    Author: Software Architect
    Date: 2024-02-29
    """
    vek = v_lx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_tx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector of tx values (number of deaths at each age).

    Parameters:
        endalter (int): Upper age limit for the calculation. If -1, use max_Alter.
        sex (str): Gender ("M" or "F").
        tafel (str): Mortality table to use.
        gebjahr (int, optional): Year of birth. Defaults to None.
        rentenbeginnalter (int, optional): Age at which pension begins. Defaults to None.
        schicht (int, optional): Pension system layer (1, 2, or 3). Defaults to 1.

    Returns:
        np.ndarray: A NumPy array containing the calculated tx values.
    """
    if endalter == -1:
        grenze = get_excel_global("max_Alter")
    else:
        grenze = endalter

    vek = np.zeros(grenze)
    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    for i in range(grenze - 1):
        vek[i] = v_temp_lx[i] - v_temp_lx[i + 1]
        vek[i] = round(vek[i], get_excel_global("rund_tx"))

    return vek

def act_tx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (tx) for a given age, sex, and mortality table.
    """
    vek = v_tx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_dx(endalter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector 'Dx' representing the present value of a life annuity due, considering various actuarial parameters.

    Parameters:
        endalter (int): The final age to which the annuity is calculated. A value of -1 indicates calculation up to 'max_Alter' (a globally defined maximum age).
        sex (str): A string indicating the sex of the annuitant. Used to select appropriate mortality rates. Example: "M" for male, "F" for female.
        tafel (str): A string specifying the mortality table to use. Different tables represent different population demographics and mortality assumptions.
        zins (float): The interest rate used for discounting future payments. This is crucial for calculating present values.
        gebjahr (int, optional): The year of birth of the annuitant. May be used in conjunction with the mortality table.
        rentenbeginnalter (int, optional): The age at which the annuity payments begin. Influences the calculation within the underlying mortality table function.
        schicht (int, optional): An integer representing the "Schicht" (layer) of the German pension system (1, 2, or 3). This likely influences the selection of specific actuarial assumptions or rules. Defaults to 1.

    Returns:
        list[float]: A list containing the calculated 'Dx' values. Each element represents the present value of a 1-Euro annuity payment due at that age, discounted at the specified interest rate and based on the provided mortality assumptions.
    """
    
    if endalter == -1:
        grenze = get_excel_global("max_Alter")
    else:
        grenze = endalter

    vek = [0.0] * grenze
    v = 1.0 / (1.0 + zins)

    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    for i in range(grenze):
        vek[i] = v_temp_lx[i] * (v ** i)
        vek[i] = round(vek[i], get_excel_global("rund_Dx"))

    return vek

cache = {}  # Initialize the cache dictionary

def act_dx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Dx' for a given age, sex, mortality table, interest rate, and pension scheme layer (Schicht).
    """
    # Überprüfe, ob das Dictionary initialisiert ist
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Dx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Überprüfen, ob der Wert bereits im Cache vorhanden ist
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Ergebnis im Cache speichern
        cache[s_key] = result
        return result

def v_cx(endalter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector of 'Cx' values, representing the present value of a life annuity-due.

    Parameters:
        endalter (int): The final age to calculate Cx for. If -1, calculations are performed up to 'max_Alter' (a globally defined constant representing the maximum age).
        sex (str): A string indicating the sex ("m" for male, "w" for female) for mortality table selection.
        tafel (str): The name of the mortality table to use.  This identifies the specific actuarial table used to determine mortality rates.
        zins (float): The interest rate used for discounting future payments.  Expressed as a decimal (e.g., 0.05 for 5%).
        gebjahr (int, optional): The year of birth. Used in conjunction with 'Tafel' to select an appropriate mortality table.  If not provided, a default table is assumed.
        rentenbeginnalter (int, optional): The age at which the annuity payments begin.  Used in conjunction with 'Tafel' and 'GebJahr'.
        schicht (int, optional): An integer representing the layer/pillar of the German pension system (1, 2, or 3).  This parameter likely influences the choice of mortality table or associated factors, reflecting different risk profiles and assumptions within each layer.  Defaults to 1.

    Returns:
        list[float]: A list containing the calculated 'Cx' values.  Each element of the list corresponds to an age.  The values are rounded to the 'rund_Cx' decimal places (a globally defined constant).
    """
    grenze = get_excel_global("max_Alter") if endalter == -1 else endalter
    vek = [0.0] * grenze
    v = 1 / (1 + zins)
    v_temp_tx = v_tx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    for i in range(grenze):
        vek[i] = v_temp_tx[i] * (v ** (i + 1))
        vek[i] = round(vek[i], get_excel_global("rund_Cx"))

    return vek

cache = {}  # Initialize the cache dictionary

def act_cx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the present value of a lifetime annuity (Actuarial Value of a Lifetime Annuity).

    Parameters:
        alter (int): The age of the insured person.
        sex (str): The gender of the insured person ("M" for male, "F" for female).
        tafel (str): The mortality table used (e.g., "DAV2018", "RSA2018").
        zins (float): The interest rate.
        gebjahr (int, optional): The year of birth of the insured person.  Required for some mortality tables.
        rentenbeginnalter (int, optional): The age at which the annuity begins.
        schicht (int, optional): The layer of pension provision (default is 1). Represents the 3-layer model of German pension provision. 1 = Statutory pension, 2 = Occupational pension/Riester, 3 = Private provision.

    Returns:
        float: The present value of the lifetime annuity.
    """
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Cx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_cx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]
        cache[s_key] = result
        return result

def v_nx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector of Nx values.

    Args:
        sex (str): Gender ("M" for male, "F" for female).
        tafel (str): Mortality table to use.
        zins (float): Discount interest rate.
        gebjahr (int, optional): Year of birth. Defaults to None.
        rentenbeginnalter (int, optional): Age at which pension payments begin. Defaults to None.
        schicht (int, optional): Layer of the pension system. Defaults to 1.

    Returns:
        list[float]: A list of Nx values for each age up to max_Alter.
    """
    vek = [0.0] * max_Alter
    v_Temp_Dx = v_dx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_Alter - 1] = v_Temp_Dx[max_Alter - 1]
    for i in range(max_Alter - 2, -1, -1):
        vek[i] = vek[i + 1] + v_Temp_Dx[i]
        vek[i] = round(vek[i], rund_Dx)

    return vek

cache = {}

def act_nx(alter: int, sex: str, tafel: str, zins: float, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial value 'Nx' representing the present value of a life annuity due, considering various actuarial and demographic parameters.
    """
    if not cache:
        initialize_cache()

    s_key = create_cache_key("Nx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_nx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        result = vek[alter]
        cache[s_key] = result
        return result

def v_mx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the ‘Mx’ vector for actuarial calculations.

    Args:
        sex (str): The sex of the annuitant ("M" or "F").
        tafel (str): The name or identifier of the mortality table.
        zins (float): The interest rate.
        gebjahr (int, optional): The year of birth. Defaults to None.
        rentenbeginnalter (int, optional): The age at which annuity payments begin. Defaults to None.
        schicht (int, optional): The layer of the German pension system. Defaults to 1.

    Returns:
        list: The ‘Mx’ vector as a list.
    """
    # Assuming max_Alter and rund_Mx are defined as global constants
    global max_Alter, rund_Mx  # Accessing global constants

    vek = [0.0] * (max_Alter + 1)

    v_Temp_Cx = v_cx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_Alter] = v_Temp_Cx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_Temp_Cx[i]
        vek[i] = round(vek[i], rund_Mx)

    return vek

cache = {}

def act_mx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Mx' for a given age, sex, mortality table, interest rate, and pension scheme layer.
    This function utilizes caching to improve performance by storing previously calculated values.

    Parameters:
        alter (int): The age for which to calculate the actuarial factor. Integer. Must be a valid age within the bounds of the mortality table.
        sex (str): The sex of the individual. String. Expected values are likely "M" for male and "F" for female. Case sensitivity may be important.
        tafel (str): The mortality table to use. String. This identifies the specific life table used for calculating probabilities of survival.
        zins (float): The interest rate to use for discounting future payments. Double. Expressed as a decimal (e.g., 0.05 for 5%).
        gebjahr (int, optional): The year of birth. Integer. May be used in conjunction with the mortality table to determine the relevant life expectancy.
        rentenbeginnalter (int, optional): The age at which the pension begins. Integer. Used for calculating the duration of payments and therefore impacts the present value factor.
        schicht (int, optional): The layer (Schicht) of the German pension system. Integer. Default is 1. This parameter likely influences the calculation based on the specific rules governing that layer (e.g., statutory pension, company pension, private pension).

    Returns:
        float: The actuarial present value factor 'Mx' as a Double. This factor is used to calculate the present value of a future pension payment.
    """
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Mx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]
        cache[s_key] = result
        return result

def v_rx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> list:
    """
    Calculates the vector of 'Rx' values, representing the present value of a life annuity due for each age.

    Parameters:
        sex (str): The sex of the annuitant ("m" for male, "w" for female). This influences mortality rates and therefore annuity calculations.
        tafel (str): The mortality table to use for calculating life expectancy. Different tables represent different population demographics and impact calculations.
        zins (float): The interest rate used for discounting future payments. Higher interest rates reduce the present value of the annuity.
        gebjahr (int, optional): The year of birth of the annuitant. Used in conjunction with 'Tafel' to determine relevant mortality rates for specific ages.
        rentenbeginnalter (int, optional): The age at which the annuity payments begin. This impacts the number of payments and their present value.
        schicht (int, optional): Indicates the layer (Schicht) of the German pension system the calculation applies to. Defaults to 1. This parameter likely influences the specific annuity assumptions or factors used (e.g., different mortality tables or risk adjustments for each layer - 1st layer = statutory pension, 2nd layer = company/Riester, 3rd layer = private).

    Returns:
        list: A list containing the calculated 'Rx' values for each age from 0 to 'max_Alter'. Each element 'vek(i)' represents the present value of an annuity-due at age 'i'. The values are rounded to 'rund_Rx' decimal places.
    """
    vek = [0.0] * (max_Alter + 1)

    v_temp_mx = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_Alter] = v_temp_mx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_mx[i]
        vek[i] = round(vek[i], rund_Rx)

    return vek

cache = {}  # Initialize the cache dictionary

def act_rx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (Rx) for a given age, sex, mortality table, interest rate, and pension scheme layer.

    Parameters:
        alter (int): The age for which to calculate the present value factor.
        sex (str): The sex of the annuitant ("M" for male, "F" for female).
        tafel (str): The mortality table to use (e.g., "DAV2018").
        zins (float): The interest rate used for the present value calculation.
        gebjahr (int, optional): The year of birth. Used in conjunction with the mortality table to determine life expectancy.
        rentenbeginnalter (int, optional): The age at which the pension begins.
        schicht (int, optional): The layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
        float: The actuarial present value factor (Rx).
    """

    # Check if the cache is initialized
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
    Calculates age based on date of birth and reference date using either
    the calendar year method or the half-year method.
    """
    if methode <> "K":
        methode = "H"

    j_gd = gebdat.year
    j_bd = berdat.year
    m_gd = gebdat.month
    m_bd = berdat.month

    if methode == "K":
        return j_bd - j_gd
    elif methode == "H":
        return int(j_bd - j_gd + 1 / 12 * (m_bd - m_gd + 5))
    else:
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
    - A string that uniquely identifies the combination of input parameters. This key is used for caching.

    Remarks:
    This function is crucial for optimizing performance within the larger pension calculation macro. 
    By creating a unique key based on all relevant input parameters, the macro can store and retrieve results from a cache, 
    avoiding the need to recalculate the same pension value repeatedly. The underscore "_" is used as a delimiter 
    to ensure a valid and easily parsable key. Understanding the German pension system's 3-Schichten-Modell is helpful 
    to interpreting the meaning of the 'Schicht' parameter.
    """
    return f"{art}_{alter}_{sex}_{tafel}_{zins}_{geb_jahr}_{rentenbeginnalter}_{schicht}"

def act_ax_k(alter: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial function 'ax_k', representing the present value of a whole life insurance benefit payable at the beginning of the year of death, adjusted by a factor 'k'.
    """
    if k > 0:
        return act_nx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_abzugsglied(k, zins)
    else:
        return 0

def act_axn_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates a factor used in actuarial calculations, specifically relating to the present value of a life annuity-due with increasing payments.
    
    Parameters:
    - alter (int): The current age of the individual.
    - n (int): The number of years for which the annuity is paid.
    - sex (str): The gender of the individual ("M" for male, "F" for female, or potentially other codes).
    - tafel (str): Identifier for the mortality table used in the calculations.
    - zins (float): The interest rate used for discounting future payments.
    - k (int): A reduction factor applied to the annuity.
    - gebjahr (int, optional): The year of birth of the individual. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin. Defaults to None.
    - schicht (int, optional): Represents a layer within the German pension system. Defaults to 1.
    
    Returns:
    - float: The calculated actuarial factor 'axn_k'. Returns 0 if 'k' is 0.
    """
    if k > 0:
        act_axn_k = (act_nx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_nx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_abzugsglied(k, zins) * (1 - act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht))
    else:
        act_axn_k = 0
    return act_axn_k

def act_nax_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates a factor used in actuarial calculations, specifically relating to the present value
    of a life annuity due to begin 'n' years from a given age.
    """
    if k > 0:
        return act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) * act_ax_k(alter + n, sex, tafel, zins, k, gebjahr, rentenbeginnalter, schicht)
    else:
        return 0

def act_ngr_ax(alter: int, n: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the average present value of a life annuity-due with a period of 'n' years.

    Parameters:
        alter (int): The current age of the individual.
        n (int): The number of years for which the annuity is paid.
        sex (str): The gender of the individual ("M" for male, "F" for female).
        tafel (str): The mortality table used for calculations.
        zins (float): The interest rate used for present value calculations.
        gebjahr (int, optional): The year of birth of the individual. Defaults to None.
        rentenbeginnalter (int, optional): The age at which the annuity payments begin. Defaults to None.
        schicht (int, optional): A layer within the German pension system. Defaults to 1.

    Returns:
        float: The calculated average present value of the annuity-due.
    """
    return (act_mx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_mx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

def act_ngr_ex(alter: int, n: int, sex: str, tafel: str, zins: float, geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the growth rate of the death probability (qx) over a period of 'n' years.

    Parameters:
        alter (int): The current age of the individual.
        n (int): The number of years into the future to calculate the growth rate.
        sex (str): The sex of the individual ("M" for male, "F" for female).
        tafel (str): The mortality table being used.
        zins (float): The interest rate used in the calculation.
        geb_jahr (int, optional): The year of birth. Defaults to None.
        rentenbeginnalter (int, optional): The age at which a pension or annuity begins. Defaults to None.
        schicht (int, optional): Indicates the layer of the German pension system. Defaults to 1.

    Returns:
        float: The growth rate of the death probability (qx) over 'n' years.
    """
    return act_dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

def act_ag_k(g: int, zins: float, k: int) -> float:
    """
    Calculates a factor used in actuarial calculations, specifically related to the present value of an increasing annuity.
    
    Parameters:
    g (int): The period or duration over which the annuity is paid.
    zins (float): The interest rate (expressed as a decimal).
    k (int): A factor influencing the calculation.
    
    Returns:
    float: The calculated factor.
    """
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
    Calculates the 'Abzugsglied' (reduction factor) used in actuarial calculations.

    Args:
        k (int): The number of years or periods.
        zins (float): The interest rate.

    Returns:
        float: The calculated 'Abzugsglied' (reduction factor).
    """
    act_abzugsglied = 0.0
    if k > 0:
        for l in range(k):
            act_abzugsglied = act_abzugsglied + l / k / (1 + l / k * zins)
        act_abzugsglied = act_abzugsglied * (1 + zins) / k
    return act_abzugsglied