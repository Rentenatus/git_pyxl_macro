

import math

import math

import math


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
        

cache = None

# -----------------------------------------------------------------------------
# Module: mConstants
# Description: This module defines constant values used throughout the pension calculation application.
#              These constants primarily control the precision and rounding behavior of actuarial calculations.
# -----------------------------------------------------------------------------

# Constant: rund_lx
# Description: Defines the rounding precision for the lifetime expectation value (lx).  This value represents the number of decimal places to which the lifetime expectation is rounded.
#              A value of 16 provides very high precision for calculations.
# Data Type: Integer
# Value: 16
rund_lx = 16

# Constant: rund_tx
# Description: Defines the rounding precision for the technical interest rate (tx). This influences the calculation of present values and future values.
#              A value of 16 ensures high precision in financial calculations.
# Data Type: Integer
# Value: 16
rund_tx = 16

# Constant: rund_Dx
# Description: Defines the rounding precision for the actuarial discount factor (Dx).  The discount factor is a key component of present value calculations.
#              A value of 16 ensures high precision in discounting future payments.
# Data Type: Integer
# Value: 16
rund_Dx = 16

# Constant: rund_Cx
# Description: Defines the rounding precision for the actuarial accumulation factor (Cx).  This factor is used to calculate the future value of an annuity or a lump sum.
#              A value of 16 provides high precision for accumulation calculations.
# Data Type: Integer
# Value: 16
rund_Cx = 16

# Constant: rund_Nx
# Description: Defines the rounding precision for the amount of a life annuity (Nx). This is the periodic payment received during a lifetime.
#              A value of 16 ensures precise calculation of annuity amounts.
# Data Type: Integer
# Value: 16
rund_Nx = 16

# Constant: rund_Mx
# Description: Defines the rounding precision for the amount of a temporary life annuity (Mx). This represents an annuity paid for a limited period.
#              A value of 16 ensures accurate calculation of temporary annuity payments.
# Data Type: Integer
# Value: 16
rund_Mx = 16

# Constant: rund_Rx
# Description: Defines the rounding precision for the remaining payments for a life annuity (Rx). This is relevant for calculating the present value of remaining annuity payments.
#              A value of 16 ensures precise calculation of present values with remaining payments.
# Data Type: Integer
# Value: 16
rund_Rx = 16

# Constant: max_Alter
# Description: Defines the maximum age considered in the pension calculations.  This establishes an upper limit for life expectancy and calculation periods.
#              A value of 123 is used as a reasonable upper limit, although actual maximum ages can vary.
# Data Type: Integer
# Value: 123
max_Alter = 123

def initialize_cache():
    # Create a new dictionary object
    global cache
    cache = {}

def act_qx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial mortality probability (qx) for a given age, sex, and mortality table.

    Parameters:
    - alter (int): The age of the individual.
    - sex (str): The sex of the individual ("M" or "F").
    - tafel (str): The mortality table identifier ("DAV1994_T" or "DAV2008_T").
    - gebjahr (int, optional): Year of birth (unused in current implementation).
    - rentenbeginnalter (int, optional): Age at pension start (unused in current implementation).
    - schicht (int, optional): Pension layer (default 1, unused in current implementation).

    Returns:
    - float: The mortality probability qx (0 to 1).

    Raises:
    - ValueError: If the specified tafel is not recognized.
    - KeyError: If required named ranges are not found.
    """
    # Normalize sex to "F" if not "M"
    sex = "F" if sex.upper() != "M" else "M"

    # Get the Tafeln worksheet
    ws = xl_workbook["Tafeln"]

    # Construct the table vector string
    sTafelvektor = tafel.upper() + "_" + sex

    # Check if the table is supported
    if tafel.upper() not in ["DAV1994_T", "DAV2008_T"]:
        raise ValueError(f"Unsupported mortality table: {tafel}")

    # Get the named ranges
    v_tafeln_range = xl_names["v_Tafeln"]
    m_tafeln_range = xl_names["m_Tafeln"]

    # Split the range references to get sheet and cell references
    v_sheet, v_ref = v_tafeln_range.split('!')
    m_sheet, m_ref = m_tafeln_range.split('!')

    # Remove quotes and dollar signs
    v_sheet = v_sheet.strip("'")
    v_ref = v_ref.replace('$', '')
    m_sheet = m_sheet.strip("'")
    m_ref = m_ref.replace('$', '')

    # Get the worksheets
    v_ws = xl_workbook[v_sheet]
    m_ws = xl_workbook[m_sheet]

    # Find the row index for the table vector
    v_range = v_ws[v_ref]
    v_values = [cell.value for cell in v_range]
    try:
        row_index = v_values.index(sTafelvektor)
    except ValueError:
        raise ValueError(f"Table vector {sTafelvektor} not found in v_Tafeln")

    # Get the mortality probability from m_Tafeln
    m_range = m_ws[m_ref]
    # Assuming m_Tafeln is a 2D range where rows are table vectors and columns are ages
    # We need to find the column that matches the alter
    # This is a simplified approach - in practice you might need to adjust based on actual data structure
    for col_idx, cell in enumerate(m_range[0], 1):  # Assuming first row is header with ages
        if cell.value == alter:
            # Get the value from the correct row and column
            return m_range[row_index][col_idx - 1].value

    raise ValueError(f"Age {alter} not found in m_Tafeln")

def v_lx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> list:
    """
    Calculates the 'lx' vector, representing the number of survivors at each age.

    Parameters:
    - endalter (int): The upper age limit for the calculation. If -1, the calculation extends to max_Alter.
    - sex (str): Gender. String value, likely "m" for male or "w" for female.
    - tafel (str): Mortality table identifier.
    - gebjahr (int, optional): Year of birth. Defaults to None.
    - rentenbeginnalter (int, optional): Age at which pension benefits begin. Defaults to None.
    - schicht (int, optional): Pension layer/pillar. Defaults to 1.

    Returns:
    - list: A list containing the 'lx' values for each age from 0 to 'endalter' (or to max_Alter if endalter is -1).
    """
    global rund_lx, max_Alter

    # Get global constants
    rund_lx = get_excel_global('rund_lx')
    max_Alter = get_excel_global('max_Alter')

    # Determine the boundary
    grenze = max_Alter if endalter == -1 else endalter

    # Initialize the vector
    vek = [0] * (grenze + 1)
    vek[0] = 1000000

    # Calculate lx values
    for i in range(1, grenze + 1):
        vek[i] = vek[i - 1] * (1 - act_qx(i - 1, sex, tafel, gebjahr, rentenbeginnalter, schicht))
        vek[i] = round(vek[i], rund_lx)

    return vek

def act_lx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (lx) for a given age, sex, and mortality table.

    Parameters:
    - alter (int): The current age for which the actuarial present value factor is calculated.
    - sex (str): Indicates the sex of the individual ("m" for male, "w" for female).
    - tafel (str): Specifies the mortality table to be used for calculations.
    - gebjahr (int, optional): The year of birth of the individual. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which pension payments are expected to begin. Defaults to None.
    - schicht (int, optional): Indicates the layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The actuarial present value factor lx for the specified age.
    """
    vek = v_lx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_tx(endalter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the 'tx' vector, representing the number of deaths at each age.
    This is a crucial component in actuarial calculations for pension and life insurance,
    specifically for determining the probability of death within a given age range.

    Parameters:
    - endalter (int): The upper age limit for the calculation. A value of -1 indicates that the calculation should extend to the maximum defined age ('max_Alter').
    - sex (str): The sex of the individual ('M' for male, 'F' for female). This influences the mortality rates used.
    - tafel (str): The mortality table to use for calculations. Different tables represent different population demographics and historical data.
    - gebjahr (int, optional): The year of birth. Used in conjunction with the mortality table to accurately reflect mortality rates for a specific cohort.
    - rentenbeginnalter (int, optional): The age at which pension payments begin. May affect the calculation depending on the specific actuarial model.
    - schicht (int, optional): An integer representing the 'Schicht' (layer) of the German three-layer pension system (1, 2, or 3). Defaults to 1 (the statutory pension layer).

    Returns:
    - list: A list containing the 'tx' values for each age from 0 up to 'endalter' (or 'max_Alter' if 'endalter' is -1). Each element represents the number of deaths at that age.
    """
    global rund_tx, max_Alter

    # Determine the boundary age
    grenze = max_Alter if endalter == -1 else endalter

    # Get the lx vector
    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    # Calculate tx values (number of deaths at each age)
    vek = []
    for i in range(grenze):
        tx_value = v_temp_lx[i] - v_temp_lx[i + 1]
        vek.append(round(tx_value, rund_tx))

    return vek

def act_tx(alter: int, sex: str, tafel: str, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor (tx) for a given age based on actuarial tables and pension scheme layer.

    Parameters:
    - alter (int): The current age of the individual.
    - sex (str): The gender of the individual ("M" for Male, "F" for Female).
    - tafel (str): The name/identifier of the actuarial table to use.
    - gebjahr (int, optional): The year of birth. Defaults to None.
    - rentenbeginnalter (int, optional): The age at which pension benefits begin. Defaults to None.
    - schicht (int, optional): The layer of the German pension system (1-3). Defaults to 1.

    Returns:
    - float: The actuarial present value factor (tx) for the given age.
    """
    vek = v_tx(alter, sex, tafel, gebjahr, rentenbeginnalter, schicht)
    return vek[alter]

def v_dx(endalter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates the vector 'Dx', representing the present value of a life annuity-due, for each age up to a specified end age.

    Parameters:
    - endalter (int): The maximum age to calculate 'Dx' for. A value of -1 indicates calculating up to the maximum supported age ('max_Alter').
    - sex (str): The sex of the individual ("M" for male, "F" for female).
    - tafel (str): The mortality table to use for calculations.
    - zins (float): The interest rate used for discounting (e.g., 0.05 for 5%).
    - gebjahr (int, optional): The year of birth.
    - rentenbeginnalter (int, optional): The age at which the annuity begins.
    - schicht (int, optional): The pension scheme layer (1, 2, or 3). Defaults to 1.

    Returns:
    - list: A list containing the 'Dx' values for each age from 0 up to 'endalter'.
    """
    global rund_Dx, max_Alter

    # Determine the boundary age
    grenze = max_Alter if endalter == -1 else endalter

    # Initialize the discount factor
    v = 1 / (1 + zins)

    # Get the lx values
    v_temp_lx = v_lx(grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    # Calculate Dx values
    vek = []
    for i in range(grenze + 1):
        dx_value = v_temp_lx[i] * (v ** i)
        vek.append(round(dx_value, rund_Dx))

    return vek

cache = None

def act_dx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
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
    Calculates the vector of 'Cx' values, representing the present value of a life annuity due.

    Parameters:
    - endalter (int): The final age to which the annuity is paid. A value of -1 indicates using the global maximum age ('max_Alter').
    - sex (str): The sex of the annuitant ("M" for male, "F" for female).
    - tafel (str): The mortality table to use for calculations.
    - zins (float): The interest rate used for discounting future payments.
    - gebjahr (int, optional): The birth year of the annuitant.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin.
    - schicht (int, optional): The layer/pillar of the German pension system (default is 1).

    Returns:
    - list[float]: A list containing the calculated 'Cx' values.
    """
    global rund_Cx, max_Alter
    rund_Cx = get_excel_global('rund_Cx')
    max_Alter = get_excel_global('max_Alter')

    Grenze = max_Alter if endalter == -1 else endalter
    vek = [0.0] * Grenze
    v = 1 / (1 + zins)

    v_temp_tx = v_tx(Grenze, sex, tafel, gebjahr, rentenbeginnalter, schicht)

    for i in range(Grenze):
        vek[i] = v_temp_tx[i] * (v ** (i + 1))
        vek[i] = round(vek[i], rund_Cx)

    return vek

cache = None

def act_cx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    global cache

    # Check if cache is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Cx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if value exists in cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_cx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store result in cache
        cache[s_key] = result
        return result

def v_nx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> list:
    """
    Calculates the vector of 'Nx' values, representing the present value of a life annuity due,
    contingent on surviving to age x.

    Parameters:
    - sex (str): The sex of the annuitant ("m" for male, "w" for female).
    - tafel (str): The mortality table to use (e.g., "DAV2018", "GUT2018").
    - zins (float): The interest rate (as a decimal, e.g., 0.05 for 5%) used for discounting.
    - gebjahr (int, optional): The year of birth of the annuitant.
    - rentenbeginnalter (int, optional): The age at which the annuity payments begin.
    - schicht (int, optional): The layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - list: A list representing the vector 'Nx'. Each element corresponds to the present value of a life annuity due contingent on surviving to that age.
    """
    global max_Alter, rund_Dx

    # Initialize the vector
    vek = [0.0] * (max_Alter + 1)  # +1 because we're using 0-based indexing

    # Get the v_Dx values
    v_temp_dx = v_dx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Set the value for max_Alter
    vek[max_Alter] = v_temp_dx[max_Alter]

    # Calculate the Nx values backwards
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_dx[i]
        vek[i] = round(vek[i], rund_Dx)

    return vek

def act_nx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Nx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value is already in cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_nx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store result in cache
        cache[s_key] = result
        return result

def v_mx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates and returns a vector (array) representing the Mx values, which are crucial in actuarial calculations
    related to German pension provisions. Mx represents the present value of a lifetime annuity, payable at the
    beginning of each period, given a certain age.

    Parameters:
    - sex (str): Indicates the sex of the individual ("M" for Male or "F" for Female)
    - tafel (str): Specifies the mortality table to be used in the calculation
    - zins (float): The interest rate used for discounting future cash flows
    - gebjahr (int, optional): The year of birth of the individual
    - rentenbeginnalter (int, optional): The age at which pension payments begin
    - schicht (int, optional): Specifies the Schicht (layer) of the German pension system (default: 1)

    Returns:
    - list[float]: An array of Mx values for each age up to max_Alter
    """
    global max_Alter, rund_Mx

    vek = [0.0] * max_Alter
    v_temp_cx = v_cx(-1, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    vek[max_Alter - 1] = v_temp_cx[max_Alter - 1]
    for i in range(max_Alter - 2, -1, -1):
        vek[i] = vek[i + 1] + v_temp_cx[i]
        vek[i] = round(vek[i], rund_Mx)

    return vek

def act_mx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    global cache

    # Check if cache is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Mx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if value is already in cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store result in cache
        cache[s_key] = result
        return result

def v_rx(sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1):
    """
    Calculates and returns a vector representing the accumulated values of Rx (present value of pension stream at each age).

    Parameters:
    - sex (str): Gender of the insured individual
    - tafel (str): Mortality table to be used
    - zins (float): Discount rate for present value calculations
    - gebjahr (int, optional): Year of birth
    - rentenbeginnalter (int, optional): Age at which pension payments begin
    - schicht (int, optional): Pension pillar (default 1)

    Returns:
    - list: Array containing accumulated Rx values at each age
    """
    global max_Alter, rund_Rx

    # Initialize the vector with size max_Alter + 1
    vek = [0.0] * (max_Alter + 1)

    # Get the v_Mx values
    v_temp_mx = v_mx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Set the value at max_Alter
    vek[max_Alter] = v_temp_mx[max_Alter]

    # Backward accumulation
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = vek[i + 1] + v_temp_mx[i]
        vek[i] = round(vek[i], rund_Rx)

    return vek

def act_rx(alter: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the actuarial present value factor 'Rx' for a given age, sex, mortality table, interest rate, and pension scheme layer.

    Parameters:
    - alter (int): The age for which to calculate the present value factor.
    - sex (str): The sex of the annuitant ("M" for male, "F" for female).
    - tafel (str): The name of the mortality table to use. This defines the probabilities of survival.
    - zins (float): The annual interest rate, expressed as a decimal.
    - gebjahr (int, optional): The year of birth. May be used in conjunction with the mortality table.
    - rentenbeginnalter (int, optional): The age at which the pension begins.
    - schicht (int, optional): The layer of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The actuarial present value factor 'Rx'. This factor represents the present value of a payment of 1 at a future date, considering mortality and discounting.

    Remarks:
    This function calculates the 'Rx' factor, a crucial component in actuarial calculations for pensions.
    It utilizes a caching mechanism to improve performance. The 'v_rx' function, which performs the core calculation,
    is called only if the result is not already present in the 'cache'.

    The 'schicht' parameter refers to the 3-layer model of German retirement provision (3-Schichten-Modell):
    - Schicht 1: Statutory pension insurance, professional pension schemes, agricultural pension funds, civil servant pensions.
    - Schicht 2: State-subsidized supplementary pensions, company pensions, Riester pensions.
    - Schicht 3: Private pension contracts without state subsidies.

    The function uses a dictionary ('cache') to store previously calculated values to avoid redundant calculations.
    """
    global cache

    # Check if the dictionary is initialized
    if cache is None:
        initialize_cache()

    s_key = create_cache_key("Rx", alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

    # Check if the value is already in the cache
    if s_key in cache:
        return cache[s_key]
    else:
        vek = v_rx(sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)
        result = vek[alter]

        # Store result in cache
        cache[s_key] = result
        return result

def act_altersberechnung(geb_dat, ber_dat, methode):
    """
    Calculates age based on date of birth and reference date using either calendar year or half-year method.

    Parameters:
    - geb_dat (date): Date of birth
    - ber_dat (date): Reference date
    - methode (str): Calculation method ("K" for calendar year, others default to half-year)

    Returns:
    - int: Calculated age in years
    """
    if methode != "K":
        methode = "H"

    j_gd = geb_dat.year
    j_bd = ber_dat.year
    m_gd = geb_dat.month
    m_bd = ber_dat.month

    if methode == "K":
        return j_bd - j_gd
    else:  # "H" method
        return int(j_bd - j_gd + 1 / 12 * (m_bd - m_gd + 5))

def create_cache_key(art: str, alter: int, sex: str, tafel: str, zins: float, geb_jahr: int, rentenbeginnalter: int, schicht: int) -> str:
    """
    Creates a unique cache key based on input parameters related to German pension calculations.

    Parameters:
    - art (str): The type of pension calculation being performed (e.g., "Sofortrente", "Dynamische Rente").
    - alter (int): The current age of the individual (in years) at the time of calculation.
    - sex (str): The gender of the individual ("m" for male, "w" for female).
    - tafel (str): The life expectancy table being used (e.g., "DAV 2018", "GUT 2018").
    - zins (float): The interest rate used in the calculation (as a decimal, e.g., 0.035 for 3.5%).
    - geb_jahr (int): The year of birth of the individual.
    - rentenbeginnalter (int): The age at which the pension is to begin (in years).
    - schicht (int): The layer (Schicht) of the German pension system being considered (1, 2, or 3).

    Returns:
    - str: A string representing a unique key for caching pension calculation results.

    Remarks:
    This function is central to a caching mechanism designed to improve performance of pension calculations.
    By creating a unique key based on the input parameters, the macro can store and retrieve previously
    calculated results, avoiding redundant computations.

    Understanding 'Schicht' (Layer):
    The German pension system is structured into three layers (3-Schichten-Modell):
    - Schicht 1: Statutory pension insurance, professional pension schemes
    - Schicht 2: State-subsidized supplementary pensions, company pension schemes
    - Schicht 3: Private pension provisions with no state subsidies
    """
    return f"{art}_{alter}_{sex}_{tafel}_{zins}_{geb_jahr}_{rentenbeginnalter}_{schicht}"

def act_ax_k(alter: int, sex: str, tafel: str, zins: float, k: int, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
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
    """
    Calculates the present value of a pension that begins in 'n' years, considering age, gender,
    mortality table, interest rate, parameter 'k' (for adjusting the present value), and optionally
    birth year, pension start age, and layer (representing type of pension provision).

    Parameters:
    - alter (int): Current age of the person
    - n (int): Number of years until pension starts
    - sex (str): Gender of the person ("m" for male, "w" for female)
    - tafel (str): Name of the mortality table used
    - zins (float): Interest rate used for discounting future payments
    - k (int): Parameter for adjusting the present value
    - gebjahr (int, optional): Birth year of the person
    - rentenbeginnalter (int, optional): Age when pension starts
    - schicht (int, optional): Layer of pension provision (default 1)

    Returns:
    - float: Present value of the pension starting in 'n' years, or 0 if k <= 0
    """
    if k > 0:
        return act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) * act_ax_k(alter + n, sex, tafel, zins, k, gebjahr, rentenbeginnalter, schicht)
    else:
        return 0.0

def act_ngr_ax(alter: int, n: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the average remaining lifetime (in years) for a cohort at a given age, considering mortality and potentially stratified by sex, table, and layer of the pension system.

    Parameters:
    - alter (int): The current age of the cohort.
    - n (int): The number of years into the future to calculate the average remaining lifetime.
    - sex (str): The sex of the cohort ("M" for male, "F" for female).
    - tafel (str): The mortality table to use for calculations.
    - zins (float): The interest rate used in present value calculations.
    - gebjahr (int, optional): The year of birth. May be relevant for selecting a mortality table based on birth cohort.
    - rentenbeginnalter (int, optional): The age at which pension payments begin. May influence calculations related to expected lifetime.
    - schicht (int, optional): Indicates the layer (Schicht) of the German pension system (1, 2, or 3). Defaults to 1.

    Returns:
    - float: The average remaining lifetime (in years) of the cohort.
    """
    return (act_mx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) - act_mx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

def act_ngr_ex(alter: int, n: int, sex: str, tafel: str, zins: float, gebjahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:
    """
    Calculates the ratio of the number of survivors at age (Alter + n) to the number of survivors at age Alter.
    This is a core component in actuarial calculations related to life insurance and pension funding.
    Specifically, it represents the probability of surviving 'n' years given survival to age 'Alter'.

    Parameters:
    - alter (int): The initial age. This is the base age for the calculation.
    - n (int): The number of years to project forward. This represents the time horizon for the survival probability calculation.
    - sex (str): The sex of the individual ("M" for male, "F" for female, or other valid codes depending on the underlying data).
    - tafel (str): The mortality table identifier. This specifies which mortality table to use for calculating survival probabilities.
    - zins (float): The interest rate. This is used for discounting future values and is fundamental in present value calculations.
    - gebjahr (int, optional): The year of birth. May be used by the underlying `act_dx` function for more refined calculations.
    - rentenbeginnalter (int, optional): The age at which a pension or benefit begins. May be used by the underlying `act_dx` function for more refined calculations.
    - schicht (int, optional): An integer representing the 'layer' or 'tier' of the pension system (default = 1). This likely relates to the 3-Schichten-Modell (3-layer model) of German pension provision.

    Returns:
    - float: The ratio of survival probabilities. Specifically, it returns act_dx(alter + n) / act_dx(alter).
    """
    return act_dx(alter + n, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht) / act_dx(alter, sex, tafel, zins, gebjahr, rentenbeginnalter, schicht)

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
    Calculates the "Abzugsglied" (deduction component) used in actuarial calculations, specifically for pension benefit calculations.

    Parameters:
    - k (int): An integer representing the number of years until the start of the pension payment.
    - zins (float): A Double representing the annual interest rate.

    Returns:
    - float: A Double representing the calculated "Abzugsglied".
    """
    abzugsglied = 0.0
    if k > 0:
        for l in range(0, k):
            abzugsglied += l / k / (1 + l / k * zins)
        abzugsglied = abzugsglied * (1 + zins) / k
    return abzugsglied