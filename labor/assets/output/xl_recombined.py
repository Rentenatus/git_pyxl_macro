

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

def fkt_kalkulation_k5():
    alpha = get_excel_global('alpha')
    beta1 = get_excel_global('beta1')
    gamma1 = get_excel_global('gamma1')
    gamma2 = get_excel_global('gamma2')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    result = (
        act_ngr_ax(x, n, Sex, Tafel, Zins) +
        act_dx(x + n, Sex, Tafel, Zins) / act_dx(x, Sex, Tafel, Zins) +
        gamma1 * act_axn_k(x, t, Sex, Tafel, Zins, 1) +
        gamma2 * (act_axn_k(x, n, Sex, Tafel, Zins, 1) - act_axn_k(x, t, Sex, Tafel, Zins, 1))
    ) / (
        (1 - beta1) * act_axn_k(x, t, Sex, Tafel, Zins, 1) - alpha * t
    )
    return result

def fkt_kalkulation_k6():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!K5")
    return VS * value

def fkt_kalkulation_k7():
    k = get_excel_global('k')
    ratzu = get_excel_global('ratzu')
    t = get_excel_global('t')
    zw = get_excel_global('zw')
    return (1 + ratzu) / zw * (get_cell_value("Kalkulation!K6") + k)

def fkt_kalkulation_k9():
    alpha = get_excel_global('alpha')
    B_xt = get_excel_global('B_xt')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    return (act_ngr_ax(x, n, Sex, Tafel, Zins) +
            act_dx(x + n, Sex, Tafel, Zins) / act_dx(x, Sex, Tafel, Zins) +
            t * alpha * B_xt) / act_axn_k(x, t, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e12():
    zw = get_excel_global('zw')
    if zw == 2:
        return 0.02
    elif zw == 4:
        return 0.03
    elif zw == 12:
        return 0.05
    else:
        return 0

def fkt_kalkulation_b16():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A16 = get_cell_value("Kalkulation!A16")

    if A16 <= n:
        return act_ngr_ax(x + A16, max(0, n - A16), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A16, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c16():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A16 = get_cell_value("Kalkulation!A16")
    return act_axn_k(x + A16, max(0, n - A16), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d16():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A16_value = get_cell_value("Kalkulation!A16")
    return act_axn_k(x + A16_value, max(0, t - A16_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e16():
    B16 = get_cell_value("Kalkulation!B16")
    D16 = get_cell_value("Kalkulation!D16")
    C16 = get_cell_value("Kalkulation!C16")
    P_xt = get_excel_global('P_xt')
    gamma2 = get_excel_global('gamma2')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B16 - P_xt * D16 + gamma2 * (C16 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D16)
    return result

def fkt_kalkulation_f16():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E16")
    return VS * value

def fkt_kalkulation_g16():
    gamma3 = get_excel_global('gamma3')
    B16 = get_cell_value("Kalkulation!B16")
    C16 = get_cell_value("Kalkulation!C16")
    return B16 + gamma3 * C16

def fkt_kalkulation_h16():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F16 = get_cell_value("Kalkulation!F16")
    A16 = get_cell_value("Kalkulation!A16")

    return F16 + alpha * t * BJB * act_axn_k(x + A16, max(5 - A16, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i16():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A16 = get_cell_value("Kalkulation!A16")
    return 1 if (x + A16 >= MinAlterFlex and A16 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j16():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A16 = get_cell_value("Kalkulation!A16")
    I16 = get_cell_value("Kalkulation!I16")
    F16 = get_cell_value("Kalkulation!F16")
    return 0 if (A16 > n) or I16 else min(150, max(50, 0.01 * (VS - F16)))

def fkt_kalkulation_k16():
    h16 = get_cell_value("Kalkulation!H16")
    j16 = get_cell_value("Kalkulation!J16")
    return max(0, h16 - j16)

def fkt_kalkulation_l16():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A16 = get_cell_value("Kalkulation!A16")
    H16 = get_cell_value("Kalkulation!H16")
    G16 = get_cell_value("Kalkulation!G16")

    try:
        if A16 > n:
            return 0
        elif A16 < t:
            return H16 / G16
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b17():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A17 = get_cell_value("Kalkulation!A17")

    if A17 <= n:
        return act_ngr_ax(x + A17, max(0, n - A17), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A17, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c17():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A17_value = get_cell_value("Kalkulation!A17")
    return act_axn_k(x + A17_value, max(0, n - A17_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d17():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A17")
    t_value = t - get_cell_value("Kalkulation!A17")
    max_value = max(0, t_value)

    return act_axn_k(alter=x_value, n=max_value, sex=Sex, tafel=Tafel, zins=Zins, k=1)

def fkt_kalkulation_e17():
    B17 = get_cell_value("Kalkulation!B17")
    P_xt = get_excel_global('P_xt')
    D17 = get_cell_value("Kalkulation!D17")
    gamma2 = get_excel_global('gamma2')
    C17 = get_cell_value("Kalkulation!C17")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B17 - P_xt * D17 + gamma2 * (C17 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D17)

def fkt_kalkulation_f17():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E17")
    return VS * value

def fkt_kalkulation_g17():
    gamma3 = get_excel_global('gamma3')
    B17 = get_cell_value("Kalkulation!B17")
    C17 = get_cell_value("Kalkulation!C17")
    return B17 + gamma3 * C17

def fkt_kalkulation_h17():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F17 = get_cell_value("Kalkulation!F17")
    A17 = get_cell_value("Kalkulation!A17")

    return F17 + alpha * t * BJB * act_axn_k(x + A17, max(5 - A17, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i17():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A17 = get_cell_value("Kalkulation!A17")
    return 1 if (x + A17 >= MinAlterFlex and A17 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j17():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A17 = get_cell_value("Kalkulation!A17")
    I17 = get_cell_value("Kalkulation!I17")
    F17 = get_cell_value("Kalkulation!F17")

    if A17 > n or I17:
        return 0
    else:
        inner_value = VS - F17
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k17():
    value = get_cell_value("Kalkulation!H17")
    value2 = get_cell_value("Kalkulation!J17")
    return max(0, value - value2)

def fkt_kalkulation_l17():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A17 = get_cell_value("Kalkulation!A17")
    H17 = get_cell_value("Kalkulation!H17")
    G17 = get_cell_value("Kalkulation!G17")
    try:
        if A17 > n:
            result = 0
        elif A17 < t:
            result = H17 / G17
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b18():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A18 = get_cell_value("Kalkulation!A18")

    if A18 <= n:
        return act_ngr_ax(x + A18, max(0, n - A18), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A18, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c18():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A18 = get_cell_value("Kalkulation!A18")
    return act_axn_k(x + A18, max(0, n - A18), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d18():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!$A18")
    t_value = t - get_cell_value("Kalkulation!$A18")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e18():
    B18 = get_cell_value("Kalkulation!B18")
    P_xt = get_excel_global('P_xt')
    D18 = get_cell_value("Kalkulation!D18")
    gamma2 = get_excel_global('gamma2')
    C18 = get_cell_value("Kalkulation!C18")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B18 - P_xt * D18 + gamma2 * (C18 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D18)

def fkt_kalkulation_f18():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E18")
    return VS * value

def fkt_kalkulation_g18():
    gamma3 = get_excel_global('gamma3')
    B18 = get_cell_value("Kalkulation!B18")
    C18 = get_cell_value("Kalkulation!C18")
    return B18 + gamma3 * C18

def fkt_kalkulation_h18():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F18 = get_cell_value("Kalkulation!F18")
    A18 = get_cell_value("Kalkulation!A18")

    return F18 + alpha * t * BJB * act_axn_k(x + A18, max(5 - A18, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i18():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A18 = get_cell_value("Kalkulation!A18")
    return 1 if (x + A18 >= MinAlterFlex and A18 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j18():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A18 = get_cell_value("Kalkulation!A18")
    I18 = get_cell_value("Kalkulation!I18")
    F18 = get_cell_value("Kalkulation!F18")

    if A18 > n or I18:
        return 0
    else:
        inner_value = VS - F18
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k18():
    H18 = get_cell_value("Kalkulation!H18")
    J18 = get_cell_value("Kalkulation!J18")
    return max(0, H18 - J18)

def fkt_kalkulation_l18():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A18 = get_cell_value("Kalkulation!A18")
    H18 = get_cell_value("Kalkulation!H18")
    G18 = get_cell_value("Kalkulation!G18")
    try:
        if A18 > n:
            return 0
        elif A18 < t:
            return H18 / G18
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b19():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A19 = get_cell_value("Kalkulation!A19")

    if A19 <= n:
        return act_ngr_ax(x + A19, max(0, n - A19), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A19, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c19():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A19_value = get_cell_value("Kalkulation!A19")
    return act_axn_k(x + A19_value, max(0, n - A19_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d19():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A19")
    t_value = t - get_cell_value("Kalkulation!A19")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e19():
    B19 = get_cell_value("Kalkulation!B19")
    P_xt = get_excel_global('P_xt')
    D19 = get_cell_value("Kalkulation!D19")
    gamma2 = get_excel_global('gamma2')
    C19 = get_cell_value("Kalkulation!C19")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B19 - P_xt * D19 + gamma2 * (C19 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D19)

def fkt_kalkulation_f19():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E19")
    return VS * value

def fkt_kalkulation_g19():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B19") + gamma3 * get_cell_value("Kalkulation!C19")
    return value

def fkt_kalkulation_h19():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F19 = get_cell_value("Kalkulation!F19")
    A19 = get_cell_value("Kalkulation!A19")

    return F19 + alpha * t * BJB * act_axn_k(x + A19, max(5 - A19, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i19():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A19 = get_cell_value("Kalkulation!A19")
    return 1 if (x + A19 >= MinAlterFlex and A19 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j19():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A19 = get_cell_value("Kalkulation!A19")
    I19 = get_cell_value("Kalkulation!I19")
    F19 = get_cell_value("Kalkulation!F19")

    if A19 > n or I19:
        return 0
    else:
        inner_value = VS - F19
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k19():
    h19 = get_cell_value("Kalkulation!H19")
    j19 = get_cell_value("Kalkulation!J19")
    return max(0, h19 - j19)

def fkt_kalkulation_l19():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A19 = get_cell_value("Kalkulation!A19")
    H19 = get_cell_value("Kalkulation!H19")
    G19 = get_cell_value("Kalkulation!G19")
    try:
        if A19 > n:
            return 0
        elif A19 < t:
            return H19 / G19
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b20():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A20 = get_cell_value("Kalkulation!A20")

    if A20 <= n:
        return act_ngr_ax(x + A20, max(0, n - A20), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A20, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c20():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    A20 = get_cell_value("Kalkulation!A20")
    return act_axn_k(x + A20, max(0, n - A20), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d20():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A20")
    t_value = t - get_cell_value("Kalkulation!A20")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e20():
    B20 = get_cell_value("Kalkulation!B20")
    P_xt = get_excel_global('P_xt')
    D20 = get_cell_value("Kalkulation!D20")
    gamma2 = get_excel_global('gamma2')
    C20 = get_cell_value("Kalkulation!C20")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B20 - P_xt * D20 + gamma2 * (C20 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D20)

def fkt_kalkulation_f20():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E20")
    return VS * value

def fkt_kalkulation_g20():
    gamma3 = get_excel_global('gamma3')
    B20 = get_cell_value("Kalkulation!B20")
    C20 = get_cell_value("Kalkulation!C20")
    return B20 + gamma3 * C20

def fkt_kalkulation_h20():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F20 = get_cell_value("Kalkulation!F20")
    A20 = get_cell_value("Kalkulation!A20")

    return F20 + alpha * t * BJB * act_axn_k(x + A20, max(5 - A20, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i20():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A20 = get_cell_value("Kalkulation!A20")
    return 1 if (x + A20 >= MinAlterFlex and A20 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j20():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A20 = get_cell_value("Kalkulation!A20")
    I20 = get_cell_value("Kalkulation!I20")
    F20 = get_cell_value("Kalkulation!F20")
    return 0 if (A20 > n) or I20 else min(150, max(50, 0.01 * (VS - F20)))

def fkt_kalkulation_k20():
    return max(0, get_cell_value("Kalkulation!H20") - get_cell_value("Kalkulation!J20"))

def fkt_kalkulation_l20():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A20 = get_cell_value("Kalkulation!A20")
    H20 = get_cell_value("Kalkulation!H20")
    G20 = get_cell_value("Kalkulation!G20")
    try:
        if A20 > n:
            result = 0
        elif A20 < t:
            result = H20 / G20
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b21():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A21 = get_cell_value("Kalkulation!A21")

    if A21 <= n:
        return act_ngr_ax(x + A21, max(0, n - A21), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A21, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c21():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!C21")
    a21_val = get_cell_value("Kalkulation!A21")
    return act_axn_k(x_val + a21_val, max(0, n - a21_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d21():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D21")
    a21_val = get_cell_value("Kalkulation!A21")
    t_val = get_cell_value("Kalkulation!A21")

    return act_axn_k(x_val + a21_val, max(0, t_val - a21_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e21():
    B21 = get_cell_value("Kalkulation!B21")
    D21 = get_cell_value("Kalkulation!D21")
    C21 = get_cell_value("Kalkulation!C21")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B21 - P_xt * D21 + gamma2 * (C21 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D21)
    return result

def fkt_kalkulation_f21():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E21")
    return VS * value

def fkt_kalkulation_g21():
    gamma3 = get_excel_global('gamma3')
    B21 = get_cell_value("Kalkulation!B21")
    C21 = get_cell_value("Kalkulation!C21")
    return B21 + gamma3 * C21

def fkt_kalkulation_h21():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F21 = get_cell_value("Kalkulation!F21")
    A21 = get_cell_value("Kalkulation!A21")

    return F21 + alpha * t * BJB * act_axn_k(x + A21, max(5 - A21, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i21():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A21 = get_cell_value("Kalkulation!A21")
    return 1 if (x + A21 >= MinAlterFlex and A21 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j21():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A21 = get_cell_value("Kalkulation!A21")
    I21 = get_cell_value("Kalkulation!I21")
    F21 = get_cell_value("Kalkulation!F21")
    return 0 if (A21 > n) or I21 else min(150, max(50, 0.01 * (VS - F21)))

def fkt_kalkulation_k21():
    h21 = get_cell_value("Kalkulation!H21")
    j21 = get_cell_value("Kalkulation!J21")
    return max(0, h21 - j21)

def fkt_kalkulation_l21():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A21 = get_cell_value("Kalkulation!A21")
    H21 = get_cell_value("Kalkulation!H21")
    G21 = get_cell_value("Kalkulation!G21")
    try:
        if A21 > n:
            result = 0
        elif A21 < t:
            result = H21 / G21
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b22():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A22 = get_cell_value("Kalkulation!A22")
    return act_ngr_ax(x + A22, max(0, n - A22), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A22, Sex, Tafel, Zins) if A22 <= n else 0

def fkt_kalkulation_c22():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A22_value = get_cell_value("Kalkulation!A22")
    return act_axn_k(x + A22_value, max(0, n - A22_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d22():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A22 = get_cell_value("Kalkulation!A22")
    return act_axn_k(x + A22, max(0, t - A22), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e22():
    B22 = get_cell_value("Kalkulation!B22")
    D22 = get_cell_value("Kalkulation!D22")
    C22 = get_cell_value("Kalkulation!C22")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B22 - P_xt * D22 + gamma2 * (C22 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D22)
    return result

def fkt_kalkulation_f22():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E22")
    return VS * value

def fkt_kalkulation_g22():
    gamma3 = get_excel_global('gamma3')
    b22 = get_cell_value("Kalkulation!B22")
    c22 = get_cell_value("Kalkulation!C22")
    return b22 + gamma3 * c22

def fkt_kalkulation_h22():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F22 = get_cell_value("Kalkulation!F22")
    A22 = get_cell_value("Kalkulation!A22")

    result = F22 + alpha * t * BJB * act_axn_k(x + A22, max(5 - A22, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)
    return result

def fkt_kalkulation_i22():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A22 = get_cell_value("Kalkulation!A22")
    return 1 if (x + A22 >= MinAlterFlex and A22 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j22():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A22 = get_cell_value("Kalkulation!A22")
    I22 = get_cell_value("Kalkulation!I22")
    F22 = get_cell_value("Kalkulation!F22")

    if A22 > n or I22:
        return 0
    else:
        result = VS - F22
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k22():
    H22 = get_cell_value("Kalkulation!H22")
    J22 = get_cell_value("Kalkulation!J22")
    return max(0, H22 - J22)

def fkt_kalkulation_l22():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A22 = get_cell_value("Kalkulation!A22")
    H22 = get_cell_value("Kalkulation!H22")
    G22 = get_cell_value("Kalkulation!G22")
    try:
        if A22 > n:
            result = 0
        else:
            if A22 < t:
                result = H22 / G22
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b23():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A23 = get_cell_value("Kalkulation!A23")
    return act_ngr_ax(x + A23, max(0, n - A23), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A23, Sex, Tafel, Zins) if A23 <= n else 0

def fkt_kalkulation_c23():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A23_value = get_cell_value("Kalkulation!A23")
    return act_axn_k(x + A23_value, max(0, n - A23_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d23():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A23")
    t_value = t - get_cell_value("Kalkulation!A23")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e23():
    B23 = get_cell_value("Kalkulation!B23")
    D23 = get_cell_value("Kalkulation!D23")
    C23 = get_cell_value("Kalkulation!C23")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B23 - P_xt * D23 + gamma2 * (C23 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D23)

def fkt_kalkulation_f23():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E23")
    return VS * value

def fkt_kalkulation_g23():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B23") + gamma3 * get_cell_value("Kalkulation!C23")
    return value

def fkt_kalkulation_h23():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F23 = get_cell_value("Kalkulation!F23")
    A23 = get_cell_value("Kalkulation!A23")

    return F23 + alpha * t * BJB * act_axn_k(x + A23, max(5 - A23, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i23():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A23 = get_cell_value("Kalkulation!A23")
    return 1 if (x + A23 >= MinAlterFlex and A23 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j23():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A23 = get_cell_value("Kalkulation!A23")
    I23 = get_cell_value("Kalkulation!I23")
    F23 = get_cell_value("Kalkulation!F23")

    if A23 > n or I23:
        return 0
    else:
        inner_value = VS - F23
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k23():
    H23 = get_cell_value("Kalkulation!H23")
    J23 = get_cell_value("Kalkulation!J23")
    return max(0, H23 - J23)

def fkt_kalkulation_l23():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A23 = get_cell_value("Kalkulation!A23")
    H23 = get_cell_value("Kalkulation!H23")
    G23 = get_cell_value("Kalkulation!G23")
    try:
        if A23 > n:
            result = 0
        elif A23 < t:
            result = H23 / G23
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b24():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A24 = get_cell_value("Kalkulation!A24")

    if A24 <= n:
        return act_ngr_ax(x + A24, max(0, n - A24), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A24, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c24():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A24_value = get_cell_value("Kalkulation!A24")
    return act_axn_k(x + A24_value, max(0, n - A24_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d24():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A24")
    t_value = t - get_cell_value("Kalkulation!A24")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e24():
    B24 = get_cell_value("Kalkulation!B24")
    D24 = get_cell_value("Kalkulation!D24")
    C24 = get_cell_value("Kalkulation!C24")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B24 - P_xt * D24 + gamma2 * (C24 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D24)
    return result

def fkt_kalkulation_f24():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E24")
    return VS * value

def fkt_kalkulation_g24():
    gamma3 = get_excel_global('gamma3')
    B24 = get_cell_value("Kalkulation!B24")
    C24 = get_cell_value("Kalkulation!C24")
    return B24 + gamma3 * C24

def fkt_kalkulation_h24():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F24 = get_cell_value("Kalkulation!F24")
    A24 = get_cell_value("Kalkulation!A24")

    return F24 + alpha * t * BJB * act_axn_k(x + A24, max(5 - A24, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i24():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A24 = get_cell_value("Kalkulation!A24")
    return 1 if (x + A24 >= MinAlterFlex and A24 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j24():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A24 = get_cell_value("Kalkulation!A24")
    I24 = get_cell_value("Kalkulation!I24")
    F24 = get_cell_value("Kalkulation!F24")

    if A24 > n or I24:
        return 0
    else:
        result = VS - F24
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k24():
    h24 = get_cell_value("Kalkulation!H24")
    j24 = get_cell_value("Kalkulation!J24")
    return max(0, h24 - j24)

def fkt_kalkulation_l24():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A24 = get_cell_value("Kalkulation!A24")
    H24 = get_cell_value("Kalkulation!H24")
    G24 = get_cell_value("Kalkulation!G24")
    try:
        if A24 > n:
            result = 0
        elif A24 < t:
            result = H24 / G24
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b25():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A25 = get_cell_value("Kalkulation!A25")

    if A25 <= n:
        return act_ngr_ax(x + A25, max(0, n - A25), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A25, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c25():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A25_value = get_cell_value("Kalkulation!A25")
    return act_axn_k(x + A25_value, max(0, n - A25_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d25():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A25")
    t_value = t - get_cell_value("Kalkulation!A25")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e25():
    B25 = get_cell_value("Kalkulation!B25")
    D25 = get_cell_value("Kalkulation!D25")
    C25 = get_cell_value("Kalkulation!C25")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B25 - P_xt * D25 + gamma2 * (C25 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D25)
    return result

def fkt_kalkulation_f25():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E25")
    return VS * value

def fkt_kalkulation_g25():
    gamma3 = get_excel_global('gamma3')
    B25 = get_cell_value("Kalkulation!B25")
    C25 = get_cell_value("Kalkulation!C25")
    return B25 + gamma3 * C25

def fkt_kalkulation_h25():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F25 = get_cell_value("Kalkulation!F25")
    A25 = get_cell_value("Kalkulation!A25")

    return F25 + alpha * t * BJB * act_axn_k(x + A25, max(5 - A25, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i25():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A25 = get_cell_value("Kalkulation!A25")
    return 1 if (x + A25 >= MinAlterFlex and A25 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j25():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A25 = get_cell_value("Kalkulation!A25")
    I25 = get_cell_value("Kalkulation!I25")
    F25 = get_cell_value("Kalkulation!F25")

    if A25 > n or I25:
        return 0
    else:
        result = VS - F25
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k25():
    h25 = get_cell_value("Kalkulation!H25")
    j25 = get_cell_value("Kalkulation!J25")
    return max(0, h25 - j25)

def fkt_kalkulation_l25():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A25 = get_cell_value("Kalkulation!A25")
    H25 = get_cell_value("Kalkulation!H25")
    G25 = get_cell_value("Kalkulation!G25")
    try:
        if A25 > n:
            return 0
        elif A25 < t:
            return H25 / G25
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b26():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A26 = get_cell_value("Kalkulation!A26")

    if A26 <= n:
        return act_ngr_ax(x + A26, max(0, n - A26), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A26, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c26():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A26 = get_cell_value("Kalkulation!A26")
    return act_axn_k(x + A26, max(0, n - A26), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d26():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = get_cell_value("Kalkulation!D26")
    a26_value = get_cell_value("Kalkulation!A26")
    return act_axn_k(x + a26_value, max(0, t - a26_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e26():
    B26 = get_cell_value("Kalkulation!B26")
    P_xt = get_excel_global('P_xt')
    D26 = get_cell_value("Kalkulation!D26")
    gamma2 = get_excel_global('gamma2')
    C26 = get_cell_value("Kalkulation!C26")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B26 - P_xt * D26 + gamma2 * (C26 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D26)

def fkt_kalkulation_f26():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E26")
    return VS * value

def fkt_kalkulation_g26():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B26") + gamma3 * get_cell_value("Kalkulation!C26")
    return value

def fkt_kalkulation_h26():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F26 = get_cell_value("Kalkulation!F26")
    A26 = get_cell_value("Kalkulation!A26")

    return F26 + alpha * t * BJB * act_axn_k(x + A26, max(5 - A26, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i26():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A26 = get_cell_value("Kalkulation!A26")
    return 1 if (x + A26 >= MinAlterFlex and A26 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j26():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A26 = get_cell_value("Kalkulation!A26")
    I26 = get_cell_value("Kalkulation!I26")
    F26 = get_cell_value("Kalkulation!F26")

    if A26 > n or I26:
        return 0
    else:
        result = VS - F26
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k26():
    value = get_cell_value("Kalkulation!H26")
    value2 = get_cell_value("Kalkulation!J26")
    return max(0, value - value2)

def fkt_kalkulation_l26():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A26 = get_cell_value("Kalkulation!A26")
    H26 = get_cell_value("Kalkulation!H26")
    G26 = get_cell_value("Kalkulation!G26")
    try:
        if A26 > n:
            return 0
        elif A26 < t:
            return H26 / G26
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b27():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A27 = get_cell_value("Kalkulation!A27")

    if A27 <= n:
        return act_ngr_ax(x + A27, max(0, n - A27), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A27, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c27():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A27 = get_cell_value("Kalkulation!A27")
    return act_axn_k(x + A27, max(0, n - A27), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d27():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A27_value = get_cell_value("Kalkulation!A27")
    return act_axn_k(x + A27_value, max(0, t - A27_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e27():
    B27 = get_cell_value("Kalkulation!B27")
    D27 = get_cell_value("Kalkulation!D27")
    C27 = get_cell_value("Kalkulation!C27")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B27 - P_xt * D27 + gamma2 * (C27 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D27)
    return result

def fkt_kalkulation_f27():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E27")
    return VS * value

def fkt_kalkulation_g27():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B27") + gamma3 * get_cell_value("Kalkulation!C27")
    return value

def fkt_kalkulation_h27():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F27 = get_cell_value("Kalkulation!F27")
    A27 = get_cell_value("Kalkulation!A27")

    return F27 + alpha * t * BJB * act_axn_k(x + A27, max(5 - A27, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i27():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A27 = get_cell_value("Kalkulation!A27")
    return 1 if (x + A27 >= MinAlterFlex and A27 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j27():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A27 = get_cell_value("Kalkulation!A27")
    I27 = get_cell_value("Kalkulation!I27")
    F27 = get_cell_value("Kalkulation!F27")

    if A27 > n or I27:
        return 0
    else:
        inner_value = VS - F27
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k27():
    h27 = get_cell_value("Kalkulation!H27")
    j27 = get_cell_value("Kalkulation!J27")
    return max(0, h27 - j27)

def fkt_kalkulation_l27():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A27 = get_cell_value("Kalkulation!A27")
    H27 = get_cell_value("Kalkulation!H27")
    G27 = get_cell_value("Kalkulation!G27")
    try:
        if A27 > n:
            result = 0
        elif A27 < t:
            result = H27 / G27
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b28():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A28 = get_cell_value("Kalkulation!A28")

    if A28 <= n:
        return act_ngr_ax(x + A28, max(0, n - A28), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A28, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c28():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A28_value = get_cell_value("Kalkulation!A28")
    return act_axn_k(x + A28_value, max(0, n - A28_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d28():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D28")
    a28_val = get_cell_value("Kalkulation!A28")
    t_val = get_cell_value("Kalkulation!A28")

    return act_axn_k(x_val + a28_val, max(0, t_val - a28_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e28():
    B28 = get_cell_value("Kalkulation!B28")
    D28 = get_cell_value("Kalkulation!D28")
    C28 = get_cell_value("Kalkulation!C28")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B28 - P_xt * D28 + gamma2 * (C28 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D28)

def fkt_kalkulation_f28():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E28")
    return VS * value

def fkt_kalkulation_g28():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B28") + gamma3 * get_cell_value("Kalkulation!C28")
    return value

def fkt_kalkulation_h28():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F28 = get_cell_value("Kalkulation!F28")
    A28 = get_cell_value("Kalkulation!A28")

    result = F28 + alpha * t * BJB * act_axn_k(x + A28, max(5 - A28, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)
    return result

def fkt_kalkulation_i28():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A28 = get_cell_value("Kalkulation!A28")
    return 1 if (x + A28 >= MinAlterFlex and A28 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j28():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A28 = get_cell_value("Kalkulation!A28")
    I28 = get_cell_value("Kalkulation!I28")
    F28 = get_cell_value("Kalkulation!F28")
    return 0 if (A28 > n) or I28 else min(150, max(50, 0.01 * (VS - F28)))

def fkt_kalkulation_k28():
    h28 = get_cell_value("Kalkulation!H28")
    j28 = get_cell_value("Kalkulation!J28")
    return max(0, h28 - j28)

def fkt_kalkulation_l28():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A28 = get_cell_value("Kalkulation!A28")
    H28 = get_cell_value("Kalkulation!H28")
    G28 = get_cell_value("Kalkulation!G28")
    try:
        if A28 > n:
            result = 0
        else:
            if A28 < t:
                result = H28 / G28
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b29():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A29 = get_cell_value("Kalkulation!A29")

    if A29 <= n:
        return act_ngr_ax(x + A29, max(0, n - A29), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A29, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c29():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A29_value = get_cell_value("Kalkulation!A29")
    return act_axn_k(x + A29_value, max(0, n - A29_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d29():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A29_value = get_cell_value("Kalkulation!A29")
    return act_axn_k(x + A29_value, max(0, t - A29_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e29():
    B29 = get_cell_value("Kalkulation!B29")
    P_xt = get_excel_global('P_xt')
    D29 = get_cell_value("Kalkulation!D29")
    gamma2 = get_excel_global('gamma2')
    C29 = get_cell_value("Kalkulation!C29")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B29 - P_xt * D29 + gamma2 * (C29 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D29)

def fkt_kalkulation_f29():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E29")
    return VS * value

def fkt_kalkulation_g29():
    gamma3 = get_excel_global('gamma3')
    B29 = get_cell_value("Kalkulation!B29")
    C29 = get_cell_value("Kalkulation!C29")
    return B29 + gamma3 * C29

def fkt_kalkulation_h29():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F29 = get_cell_value("Kalkulation!F29")
    A29 = get_cell_value("Kalkulation!A29")

    return F29 + alpha * t * BJB * act_axn_k(x + A29, max(5 - A29, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i29():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A29 = get_cell_value("Kalkulation!A29")
    return 1 if (x + A29 >= MinAlterFlex and A29 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j29():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A29 = get_cell_value("Kalkulation!A29")
    I29 = get_cell_value("Kalkulation!I29")
    F29 = get_cell_value("Kalkulation!F29")

    if A29 > n or I29:
        return 0
    else:
        result = VS - F29
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k29():
    H29 = get_cell_value("Kalkulation!H29")
    J29 = get_cell_value("Kalkulation!J29")
    return max(0, H29 - J29)

def fkt_kalkulation_l29():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A29 = get_cell_value("Kalkulation!A29")
    H29 = get_cell_value("Kalkulation!H29")
    G29 = get_cell_value("Kalkulation!G29")
    try:
        if A29 > n:
            result = 0
        elif A29 < t:
            result = H29 / G29
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b30():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A30 = get_cell_value("Kalkulation!A30")

    if A30 <= n:
        return act_ngr_ax(x + A30, max(0, n - A30), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A30, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c30():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A30_value = get_cell_value("Kalkulation!A30")
    return act_axn_k(x + A30_value, max(0, n - A30_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d30():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D30")
    a30_val = get_cell_value("Kalkulation!A30")
    t_val = get_cell_value("Kalkulation!A30")

    return act_axn_k(x_val + a30_val, max(0, t_val - a30_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e30():
    B30 = get_cell_value("Kalkulation!B30")
    D30 = get_cell_value("Kalkulation!D30")
    C30 = get_cell_value("Kalkulation!C30")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B30 - P_xt * D30 + gamma2 * (C30 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D30)

def fkt_kalkulation_f30():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E30")
    return VS * value

def fkt_kalkulation_g30():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B30") + gamma3 * get_cell_value("Kalkulation!C30")
    return value

def fkt_kalkulation_h30():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F30 = get_cell_value("Kalkulation!F30")
    A30 = get_cell_value("Kalkulation!A30")

    return F30 + alpha * t * BJB * act_axn_k(x + A30, max(5 - A30, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i30():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A30 = get_cell_value("Kalkulation!A30")
    return 1 if (x + A30 >= MinAlterFlex and A30 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j30():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A30 = get_cell_value("Kalkulation!A30")
    I30 = get_cell_value("Kalkulation!I30")
    F30 = get_cell_value("Kalkulation!F30")
    return 0 if (A30 > n) or I30 else min(150, max(50, 0.01 * (VS - F30)))

def fkt_kalkulation_k30():
    H30 = get_cell_value("Kalkulation!H30")
    J30 = get_cell_value("Kalkulation!J30")
    return max(0, H30 - J30)

def fkt_kalkulation_l30():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A30 = get_cell_value("Kalkulation!A30")
    H30 = get_cell_value("Kalkulation!H30")
    G30 = get_cell_value("Kalkulation!G30")
    try:
        if A30 > n:
            result = 0
        elif A30 < t:
            result = H30 / G30
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b31():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A31 = get_cell_value("Kalkulation!A31")
    result = 0
    if A31 <= n:
        term1 = act_ngr_ax(x + A31, max(0, n - A31), Sex, Tafel, Zins)
        term2 = act_dx(x + n, Sex, Tafel, Zins)
        term3 = act_dx(x + A31, Sex, Tafel, Zins)
        result = term1 + (term2 / term3)
    return result

def fkt_kalkulation_c31():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!C31")
    a31_val = get_cell_value("Kalkulation!A31")
    return act_axn_k(x_val + a31_val, max(0, n - a31_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d31():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A31")
    t_value = t - get_cell_value("Kalkulation!A31")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e31():
    B31 = get_cell_value("Kalkulation!B31")
    P_xt = get_excel_global('P_xt')
    D31 = get_cell_value("Kalkulation!D31")
    gamma2 = get_excel_global('gamma2')
    C31 = get_cell_value("Kalkulation!C31")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B31 - P_xt * D31 + gamma2 * (C31 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D31)

def fkt_kalkulation_f31():
    VS = get_excel_global('VS')
    E31 = get_cell_value("Kalkulation!E31")
    return VS * E31

def fkt_kalkulation_g31():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B31") + gamma3 * get_cell_value("Kalkulation!C31")
    return value

def fkt_kalkulation_h31():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F31 = get_cell_value("Kalkulation!F31")
    A31 = get_cell_value("Kalkulation!A31")

    return F31 + alpha * t * BJB * act_axn_k(x + A31, max(5 - A31, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i31():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A31 = get_cell_value("Kalkulation!A31")
    return 1 if (x + A31 >= MinAlterFlex and A31 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j31():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A31 = get_cell_value("Kalkulation!A31")
    I31 = get_cell_value("Kalkulation!I31")
    F31 = get_cell_value("Kalkulation!F31")

    if A31 > n or I31:
        return 0
    else:
        inner_value = VS - F31
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k31():
    value = get_cell_value("Kalkulation!H31")
    value2 = get_cell_value("Kalkulation!J31")
    return max(0, value - value2)

def fkt_kalkulation_l31():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A31 = get_cell_value("Kalkulation!A31")
    H31 = get_cell_value("Kalkulation!H31")
    G31 = get_cell_value("Kalkulation!G31")
    try:
        if A31 > n:
            return 0
        elif A31 < t:
            return H31 / G31
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b32():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A32 = get_cell_value("Kalkulation!A32")

    if A32 <= n:
        return act_ngr_ax(x + A32, max(0, n - A32), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A32, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c32():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A32_value = get_cell_value("Kalkulation!A32")
    return act_axn_k(x + A32_value, max(0, n - A32_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d32():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D32")
    a32_val = get_cell_value("Kalkulation!A32")
    t_val = get_cell_value("Kalkulation!A32")

    return act_axn_k(x_val + a32_val, max(0, t_val - a32_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e32():
    B32 = get_cell_value("Kalkulation!B32")
    D32 = get_cell_value("Kalkulation!D32")
    C32 = get_cell_value("Kalkulation!C32")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B32 - P_xt * D32 + gamma2 * (C32 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D32)
    return result

def fkt_kalkulation_f32():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E32")
    return VS * value

def fkt_kalkulation_g32():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B32") + gamma3 * get_cell_value("Kalkulation!C32")
    return value

def fkt_kalkulation_h32():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F32 = get_cell_value("Kalkulation!F32")
    A32 = get_cell_value("Kalkulation!A32")

    return F32 + alpha * t * BJB * act_axn_k(x + A32, max(5 - A32, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i32():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A32 = get_cell_value("Kalkulation!A32")
    return 1 if (x + A32 >= MinAlterFlex and A32 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j32():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A32_value = get_cell_value("Kalkulation!A32")
    I32_value = get_cell_value("Kalkulation!I32")
    F32_value = get_cell_value("Kalkulation!F32")

    if A32_value > n or I32_value:
        return 0
    else:
        result = VS - F32_value
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k32():
    H32 = get_cell_value("Kalkulation!H32")
    J32 = get_cell_value("Kalkulation!J32")
    return max(0, H32 - J32)

def fkt_kalkulation_l32():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A32 = get_cell_value("Kalkulation!A32")
    H32 = get_cell_value("Kalkulation!H32")
    G32 = get_cell_value("Kalkulation!G32")
    try:
        if A32 > n:
            result = 0
        elif A32 < t:
            result = H32 / G32
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b33():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A33 = get_cell_value("Kalkulation!A33")

    if A33 <= n:
        return act_ngr_ax(x + A33, max(0, n - A33), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A33, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c33():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A33_value = get_cell_value("Kalkulation!A33")
    return act_axn_k(x + A33_value, max(0, n - A33_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d33():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D33")
    a33_val = get_cell_value("Kalkulation!A33")
    t_val = get_cell_value("Kalkulation!A33")

    return act_axn_k(x_val + a33_val, max(0, t_val - a33_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e33():
    B33 = get_cell_value("Kalkulation!B33")
    P_xt = get_excel_global('P_xt')
    D33 = get_cell_value("Kalkulation!D33")
    gamma2 = get_excel_global('gamma2')
    C33 = get_cell_value("Kalkulation!C33")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B33 - P_xt * D33 + gamma2 * (C33 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D33)
    return result

def fkt_kalkulation_f33():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E33")
    return VS * value

def fkt_kalkulation_g33():
    gamma3 = get_excel_global('gamma3')
    B33 = get_cell_value("Kalkulation!B33")
    C33 = get_cell_value("Kalkulation!C33")
    return B33 + gamma3 * C33

def fkt_kalkulation_h33():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F33 = get_cell_value("Kalkulation!F33")
    A33 = get_cell_value("Kalkulation!A33")

    return F33 + alpha * t * BJB * act_axn_k(x + A33, max(5 - A33, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i33():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A33 = get_cell_value("Kalkulation!A33")
    return 1 if (x + A33 >= MinAlterFlex and A33 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j33():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A33 = get_cell_value("Kalkulation!A33")
    I33 = get_cell_value("Kalkulation!I33")
    F33 = get_cell_value("Kalkulation!F33")

    if A33 > n or I33:
        return 0
    else:
        result = VS - F33
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k33():
    h33 = get_cell_value("Kalkulation!H33")
    j33 = get_cell_value("Kalkulation!J33")
    return max(0, h33 - j33)

def fkt_kalkulation_l33():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A33 = get_cell_value("Kalkulation!A33")
    H33 = get_cell_value("Kalkulation!H33")
    G33 = get_cell_value("Kalkulation!G33")
    try:
        if A33 > n:
            return 0
        elif A33 < t:
            return H33 / G33
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b34():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A34 = get_cell_value("Kalkulation!A34")

    if A34 <= n:
        return act_ngr_ax(x + A34, max(0, n - A34), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A34, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c34():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A34_value = get_cell_value("Kalkulation!A34")
    return act_axn_k(x + A34_value, max(0, n - A34_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d34():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A34")
    t_value = t - get_cell_value("Kalkulation!A34")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e34():
    B34 = get_cell_value("Kalkulation!B34")
    D34 = get_cell_value("Kalkulation!D34")
    C34 = get_cell_value("Kalkulation!C34")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B34 - P_xt * D34 + gamma2 * (C34 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D34)

def fkt_kalkulation_f34():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E34")
    return VS * value

def fkt_kalkulation_g34():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B34") + gamma3 * get_cell_value("Kalkulation!C34")
    return value

def fkt_kalkulation_h34():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F34 = get_cell_value("Kalkulation!F34")
    A34 = get_cell_value("Kalkulation!A34")

    return F34 + alpha * t * BJB * act_axn_k(x + A34, max(5 - A34, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i34():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A34 = get_cell_value("Kalkulation!A34")
    return 1 if (x + A34 >= MinAlterFlex and A34 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j34():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A34 = get_cell_value("Kalkulation!A34")
    I34 = get_cell_value("Kalkulation!I34")
    F34 = get_cell_value("Kalkulation!F34")

    if A34 > n or I34:
        return 0
    else:
        inner_value = VS - F34
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k34():
    value = get_cell_value("Kalkulation!H34")
    value2 = get_cell_value("Kalkulation!J34")
    return max(0, value - value2)

def fkt_kalkulation_l34():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A34 = get_cell_value("Kalkulation!A34")
    H34 = get_cell_value("Kalkulation!H34")
    G34 = get_cell_value("Kalkulation!G34")
    try:
        if A34 > n:
            return 0
        elif A34 < t:
            return H34 / G34
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b35():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A35 = get_cell_value("Kalkulation!A35")
    result = act_ngr_ax(x + A35, max(0, n - A35), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A35, Sex, Tafel, Zins) if A35 <= n else 0
    return result

def fkt_kalkulation_c35():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!C35")
    a35_val = get_cell_value("Kalkulation!A35")
    return act_axn_k(x_val + a35_val, max(0, n - a35_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d35():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D35")
    a35_val = get_cell_value("Kalkulation!A35")
    t_val = get_cell_value("Kalkulation!A35")

    return act_axn_k(x_val + a35_val, max(0, t_val - a35_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e35():
    B35 = get_cell_value("Kalkulation!B35")
    D35 = get_cell_value("Kalkulation!D35")
    C35 = get_cell_value("Kalkulation!C35")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')
    P_xt = get_excel_global('P_xt')
    gamma2 = get_excel_global('gamma2')

    return B35 - P_xt * D35 + gamma2 * (C35 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D35)

def fkt_kalkulation_f35():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E35")
    return VS * value

def fkt_kalkulation_g35():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B35") + gamma3 * get_cell_value("Kalkulation!C35")
    return value

def fkt_kalkulation_h35():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F35 = get_cell_value("Kalkulation!F35")
    A35 = get_cell_value("Kalkulation!A35")

    return F35 + alpha * t * BJB * act_axn_k(x + A35, max(5 - A35, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i35():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A35 = get_cell_value("Kalkulation!A35")
    return 1 if (x + A35 >= MinAlterFlex and A35 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j35():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A35 = get_cell_value("Kalkulation!A35")
    I35 = get_cell_value("Kalkulation!I35")
    F35 = get_cell_value("Kalkulation!F35")

    if A35 > n or I35:
        return 0
    else:
        result = VS - F35
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k35():
    H35 = get_cell_value("Kalkulation!H35")
    J35 = get_cell_value("Kalkulation!J35")
    return max(0, H35 - J35)

def fkt_kalkulation_l35():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A35 = get_cell_value("Kalkulation!A35")
    H35 = get_cell_value("Kalkulation!H35")
    G35 = get_cell_value("Kalkulation!G35")
    try:
        if A35 > n:
            return 0
        elif A35 < t:
            return H35 / G35
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b36():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A36 = get_cell_value("Kalkulation!A36")
    return act_ngr_ax(x + A36, max(0, n - A36), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A36, Sex, Tafel, Zins) if A36 <= n else 0

def fkt_kalkulation_c36():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A36_value = get_cell_value("Kalkulation!A36")
    return act_axn_k(x + A36_value, max(0, n - A36_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d36():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A36")
    t_value = t - get_cell_value("Kalkulation!A36")
    max_value = max(0, t_value)

    return act_axn_k(alter=x_value, n=max_value, sex=Sex, tafel=Tafel, zins=Zins, k=1)

def fkt_kalkulation_e36():
    B36 = get_cell_value("Kalkulation!B36")
    D36 = get_cell_value("Kalkulation!D36")
    C36 = get_cell_value("Kalkulation!C36")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B36 - P_xt * D36 + gamma2 * (C36 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D36)

def fkt_kalkulation_f36():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E36")
    return VS * value

def fkt_kalkulation_g36():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B36") + gamma3 * get_cell_value("Kalkulation!C36")
    return value

def fkt_kalkulation_h36():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F36 = get_cell_value("Kalkulation!F36")
    A36 = get_cell_value("Kalkulation!A36")

    return F36 + alpha * t * BJB * act_axn_k(x + A36, max(5 - A36, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i36():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A36 = get_cell_value("Kalkulation!A36")
    return 1 if (x + A36 >= MinAlterFlex and A36 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j36():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A36 = get_cell_value("Kalkulation!A36")
    I36 = get_cell_value("Kalkulation!I36")
    F36 = get_cell_value("Kalkulation!F36")
    if A36 > n or I36:
        return 0
    else:
        inner_value = VS - F36
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k36():
    H36 = get_cell_value("Kalkulation!H36")
    J36 = get_cell_value("Kalkulation!J36")
    return max(0, H36 - J36)

def fkt_kalkulation_l36():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A36 = get_cell_value("Kalkulation!A36")
    H36 = get_cell_value("Kalkulation!H36")
    G36 = get_cell_value("Kalkulation!G36")
    try:
        if A36 > n:
            result = 0
        else:
            if A36 < t:
                result = H36 / G36
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b37():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A37 = get_cell_value("Kalkulation!A37")
    if A37 <= n:
        return act_ngr_ax(x + A37, max(0, n - A37), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A37, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c37():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A37_value = get_cell_value("Kalkulation!A37")
    return act_axn_k(x + A37_value, max(0, n - A37_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d37():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A37_value = get_cell_value("Kalkulation!A37")
    return act_axn_k(x + A37_value, max(0, t - A37_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e37():
    B37 = get_cell_value("Kalkulation!B37")
    P_xt = get_excel_global('P_xt')
    D37 = get_cell_value("Kalkulation!D37")
    gamma2 = get_excel_global('gamma2')
    C37 = get_cell_value("Kalkulation!C37")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B37 - P_xt * D37 + gamma2 * (C37 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D37)

def fkt_kalkulation_f37():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E37")
    return VS * value

def fkt_kalkulation_g37():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B37") + gamma3 * get_cell_value("Kalkulation!C37")
    return value

def fkt_kalkulation_h37():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F37 = get_cell_value("Kalkulation!F37")
    A37 = get_cell_value("Kalkulation!A37")

    result = F37 + alpha * t * BJB * act_axn_k(x + A37, max(5 - A37, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)
    return result

def fkt_kalkulation_i37():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A37 = get_cell_value("Kalkulation!A37")
    return 1 if (x + A37 >= MinAlterFlex and A37 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j37():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A37 = get_cell_value("Kalkulation!A37")
    I37 = get_cell_value("Kalkulation!I37")
    F37 = get_cell_value("Kalkulation!F37")

    if A37 > n or I37:
        return 0
    else:
        inner_value = VS - F37
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k37():
    H37 = get_cell_value("Kalkulation!H37")
    J37 = get_cell_value("Kalkulation!J37")
    return max(0, H37 - J37)

def fkt_kalkulation_l37():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A37 = get_cell_value("Kalkulation!A37")
    H37 = get_cell_value("Kalkulation!H37")
    G37 = get_cell_value("Kalkulation!G37")
    try:
        if A37 > n:
            result = 0
        else:
            if A37 < t:
                result = H37 / G37
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b38():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A38 = get_cell_value("Kalkulation!A38")

    if A38 <= n:
        return act_ngr_ax(x + A38, max(0, n - A38), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A38, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c38():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A38_value = get_cell_value("Kalkulation!A38")
    return act_axn_k(x + A38_value, max(0, n - A38_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d38():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A38_value = get_cell_value("Kalkulation!A38")
    return act_axn_k(x + A38_value, max(0, t - A38_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e38():
    B38 = get_cell_value("Kalkulation!B38")
    D38 = get_cell_value("Kalkulation!D38")
    C38 = get_cell_value("Kalkulation!C38")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    term1 = B38 - P_xt * D38
    term2 = gamma2 * (C38 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D38)
    return term1 + term2

def fkt_kalkulation_f38():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E38")
    return VS * value

def fkt_kalkulation_g38():
    gamma3 = get_excel_global('gamma3')
    b38 = get_cell_value("Kalkulation!B38")
    c38 = get_cell_value("Kalkulation!C38")
    return b38 + gamma3 * c38

def fkt_kalkulation_h38():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F38 = get_cell_value("Kalkulation!F38")
    A38 = get_cell_value("Kalkulation!A38")

    return F38 + alpha * t * BJB * act_axn_k(x + A38, max(5 - A38, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i38():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A38 = get_cell_value("Kalkulation!A38")
    return 1 if (x + A38 >= MinAlterFlex and A38 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j38():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A38 = get_cell_value("Kalkulation!A38")
    I38 = get_cell_value("Kalkulation!I38")
    F38 = get_cell_value("Kalkulation!F38")

    if A38 > n or I38:
        return 0
    else:
        inner_value = VS - F38
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k38():
    H38 = get_cell_value("Kalkulation!H38")
    J38 = get_cell_value("Kalkulation!J38")
    return max(0, H38 - J38)

def fkt_kalkulation_l38():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A38 = get_cell_value("Kalkulation!A38")
    H38 = get_cell_value("Kalkulation!H38")
    G38 = get_cell_value("Kalkulation!G38")
    try:
        if A38 > n:
            return 0
        elif A38 < t:
            return H38 / G38
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b39():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A39 = get_cell_value("Kalkulation!A39")
    return act_ngr_ax(x + A39, max(0, n - A39), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A39, Sex, Tafel, Zins) if A39 <= n else 0

def fkt_kalkulation_c39():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A39_value = get_cell_value("Kalkulation!A39")
    return act_axn_k(x + A39_value, max(0, n - A39_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d39():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A39_value = get_cell_value("Kalkulation!A39")
    return act_axn_k(x + A39_value, max(0, t - A39_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e39():
    gamma2 = get_excel_global('gamma2')
    k = get_excel_global('k')
    n = get_excel_global('n')
    P_xt = get_excel_global('P_xt')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    B39 = get_cell_value("Kalkulation!B39")
    D39 = get_cell_value("Kalkulation!D39")
    C39 = get_cell_value("Kalkulation!C39")

    result = B39 - P_xt * D39 + gamma2 * (C39 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D39)
    return result

def fkt_kalkulation_f39():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E39")
    return VS * value

def fkt_kalkulation_g39():
    gamma3 = get_excel_global('gamma3')
    b39 = get_cell_value("Kalkulation!B39")
    c39 = get_cell_value("Kalkulation!C39")
    return b39 + gamma3 * c39

def fkt_kalkulation_h39():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F39 = get_cell_value("Kalkulation!F39")
    A39 = get_cell_value("Kalkulation!A39")

    return F39 + alpha * t * BJB * act_axn_k(x + A39, max(5 - A39, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i39():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A39 = get_cell_value("Kalkulation!A39")
    return 1 if (x + A39 >= MinAlterFlex and A39 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j39():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A39 = get_cell_value("Kalkulation!A39")
    I39 = get_cell_value("Kalkulation!I39")
    F39 = get_cell_value("Kalkulation!F39")

    if A39 > n or I39:
        return 0
    else:
        inner_value = VS - F39
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k39():
    H39 = get_cell_value("Kalkulation!H39")
    J39 = get_cell_value("Kalkulation!J39")
    return max(0, H39 - J39)

def fkt_kalkulation_l39():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A39 = get_cell_value("Kalkulation!A39")
    H39 = get_cell_value("Kalkulation!H39")
    G39 = get_cell_value("Kalkulation!G39")
    try:
        if A39 > n:
            result = 0
        else:
            if A39 < t:
                result = H39 / G39
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b40():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A40 = get_cell_value("Kalkulation!A40")
    return act_ngr_ax(x + A40, max(0, n - A40), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A40, Sex, Tafel, Zins) if A40 <= n else 0

def fkt_kalkulation_c40():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A40_value = get_cell_value("Kalkulation!A40")
    return act_axn_k(x + A40_value, max(0, n - A40_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d40():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A40")
    t_value = t - get_cell_value("Kalkulation!A40")
    max_value = max(0, t_value)

    return act_axn_k(alter=x_value, n=max_value, sex=Sex, tafel=Tafel, zins=Zins, k=1)

def fkt_kalkulation_e40():
    gamma2 = get_excel_global('gamma2')
    k = get_excel_global('k')
    n = get_excel_global('n')
    P_xt = get_excel_global('P_xt')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    B40 = get_cell_value("Kalkulation!B40")
    D40 = get_cell_value("Kalkulation!D40")
    C40 = get_cell_value("Kalkulation!C40")

    return B40 - P_xt * D40 + gamma2 * (C40 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D40)

def fkt_kalkulation_f40():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E40")
    return VS * value

def fkt_kalkulation_g40():
    gamma3 = get_excel_global('gamma3')
    B40 = get_cell_value("Kalkulation!B40")
    C40 = get_cell_value("Kalkulation!C40")
    return B40 + gamma3 * C40

def fkt_kalkulation_h40():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F40 = get_cell_value("Kalkulation!F40")
    A40 = get_cell_value("Kalkulation!A40")

    return F40 + alpha * t * BJB * act_axn_k(x + A40, max(5 - A40, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i40():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A40 = get_cell_value("Kalkulation!A40")
    return 1 if (x + A40 >= MinAlterFlex and A40 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j40():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A40 = get_cell_value("Kalkulation!A40")
    I40 = get_cell_value("Kalkulation!I40")
    F40 = get_cell_value("Kalkulation!F40")

    if A40 > n or I40:
        return 0
    else:
        inner_value = VS - F40
        result = inner_value * 0.01
        return min(150, max(50, result))

def fkt_kalkulation_k40():
    value = get_cell_value("Kalkulation!H40")
    value2 = get_cell_value("Kalkulation!J40")
    return max(0, value - value2)

def fkt_kalkulation_l40():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A40 = get_cell_value("Kalkulation!A40")
    H40 = get_cell_value("Kalkulation!H40")
    G40 = get_cell_value("Kalkulation!G40")
    try:
        if A40 > n:
            result = 0
        elif A40 < t:
            result = H40 / G40
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b41():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A41 = get_cell_value("Kalkulation!A41")

    if A41 <= n:
        return act_ngr_ax(x + A41, max(0, n - A41), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A41, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c41():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A41_value = get_cell_value("Kalkulation!A41")
    return act_axn_k(x + A41_value, max(0, n - A41_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d41():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A41 = get_cell_value("Kalkulation!A41")
    return act_axn_k(x + A41, max(0, t - A41), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e41():
    gamma2 = get_excel_global('gamma2')
    k = get_excel_global('k')
    n = get_excel_global('n')
    P_xt = get_excel_global('P_xt')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    B41 = get_cell_value("Kalkulation!B41")
    D41 = get_cell_value("Kalkulation!D41")
    C41 = get_cell_value("Kalkulation!C41")

    return B41 - P_xt * D41 + gamma2 * (C41 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D41)

def fkt_kalkulation_f41():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E41")
    return VS * value

def fkt_kalkulation_g41():
    gamma3 = get_excel_global('gamma3')
    return get_cell_value("Kalkulation!B41") + gamma3 * get_cell_value("Kalkulation!C41")

def fkt_kalkulation_h41() -> float:
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F41 = get_cell_value("Kalkulation!F41")
    A41 = get_cell_value("Kalkulation!A41")

    return F41 + alpha * t * BJB * act_axn_k(x + A41, max(5 - A41, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i41():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A41 = get_cell_value("Kalkulation!A41")
    return 1 if (x + A41 >= MinAlterFlex and A41 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j41():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A41 = get_cell_value("Kalkulation!A41")
    I41 = get_cell_value("Kalkulation!I41")
    F41 = get_cell_value("Kalkulation!F41")

    if A41 > n or I41:
        return 0
    else:
        inner_value = VS - F41
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k41():
    h41 = get_cell_value("Kalkulation!H41")
    j41 = get_cell_value("Kalkulation!J41")
    return max(0, h41 - j41)

def fkt_kalkulation_l41():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A41 = get_cell_value("Kalkulation!A41")
    H41 = get_cell_value("Kalkulation!H41")
    G41 = get_cell_value("Kalkulation!G41")
    try:
        if A41 > n:
            result = 0
        else:
            if A41 < t:
                result = H41 / G41
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b42():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A42 = get_cell_value("Kalkulation!A42")

    if A42 <= n:
        return act_ngr_ax(x + A42, max(0, n - A42), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A42, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c42():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A42_value = get_cell_value("Kalkulation!A42")
    return act_axn_k(x + A42_value, max(0, n - A42_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d42():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D42")
    A42_val = get_cell_value("Kalkulation!A42")
    t_val = get_cell_value("Kalkulation!A42")

    return act_axn_k(x_val + A42_val, max(0, t_val - A42_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e42():
    B42 = get_cell_value("Kalkulation!B42")
    D42 = get_cell_value("Kalkulation!D42")
    C42 = get_cell_value("Kalkulation!C42")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')
    P_xt = get_excel_global('P_xt')
    gamma2 = get_excel_global('gamma2')

    return B42 - P_xt * D42 + gamma2 * (C42 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D42)

def fkt_kalkulation_f42():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E42")
    return VS * value

def fkt_kalkulation_g42():
    gamma3 = get_excel_global('gamma3')
    B42 = get_cell_value("Kalkulation!B42")
    C42 = get_cell_value("Kalkulation!C42")
    return B42 + gamma3 * C42

def fkt_kalkulation_h42():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F42 = get_cell_value("Kalkulation!F42")
    A42 = get_cell_value("Kalkulation!A42")

    return F42 + alpha * t * BJB * act_axn_k(x + A42, max(5 - A42, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i42():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A42 = get_cell_value("Kalkulation!A42")
    return 1 if (x + A42 >= MinAlterFlex and A42 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j42():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A42 = get_cell_value("Kalkulation!A42")
    I42 = get_cell_value("Kalkulation!I42")
    F42 = get_cell_value("Kalkulation!F42")

    if A42 > n or I42:
        return 0
    else:
        inner_value = VS - F42
        result = inner_value * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k42():
    value = get_cell_value("Kalkulation!H42")
    value2 = get_cell_value("Kalkulation!J42")
    return max(0, value - value2)

def fkt_kalkulation_l42():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A42 = get_cell_value("Kalkulation!A42")
    H42 = get_cell_value("Kalkulation!H42")
    G42 = get_cell_value("Kalkulation!G42")
    try:
        if A42 > n:
            result = 0
        else:
            if A42 < t:
                result = H42 / G42
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b43():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A43 = get_cell_value("Kalkulation!A43")

    if A43 <= n:
        return act_ngr_ax(x + A43, max(0, n - A43), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A43, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c43():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A43_value = get_cell_value("Kalkulation!A43")
    return act_axn_k(x + A43_value, max(0, n - A43_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d43():
    x = get_excel_global('x')
    t = get_excel_global('t')
    A43 = get_cell_value("Kalkulation!A43")
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    return act_axn_k(x + A43, max(0, t - A43), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e43():
    B43 = get_cell_value("Kalkulation!B43")
    D43 = get_cell_value("Kalkulation!D43")
    C43 = get_cell_value("Kalkulation!C43")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B43 - P_xt * D43 + gamma2 * (C43 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D43)

def fkt_kalkulation_f43():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E43")
    return VS * value

def fkt_kalkulation_g43():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B43") + gamma3 * get_cell_value("Kalkulation!C43")
    return value

def fkt_kalkulation_h43():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F43 = get_cell_value("Kalkulation!F43")
    A43 = get_cell_value("Kalkulation!A43")

    return F43 + alpha * t * BJB * act_axn_k(x + A43, max(5 - A43, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i43():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A43 = get_cell_value("Kalkulation!A43")
    return 1 if (x + A43 >= MinAlterFlex and A43 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j43():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A43_value = get_cell_value("Kalkulation!A43")
    I43_value = get_cell_value("Kalkulation!I43")
    F43_value = get_cell_value("Kalkulation!F43")

    if A43_value > n or I43_value:
        return 0
    else:
        inner_value = VS - F43_value
        result = inner_value * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k43():
    h43 = get_cell_value("Kalkulation!H43")
    j43 = get_cell_value("Kalkulation!J43")
    return max(0, h43 - j43)

def fkt_kalkulation_l43():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A43 = get_cell_value("Kalkulation!A43")
    H43 = get_cell_value("Kalkulation!H43")
    G43 = get_cell_value("Kalkulation!G43")
    try:
        if A43 > n:
            return 0
        elif A43 < t:
            return H43 / G43
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b44():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A44 = get_cell_value("Kalkulation!A44")

    if A44 <= n:
        return act_ngr_ax(x + A44, max(0, n - A44), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A44, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c44():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!C44")
    a44_val = get_cell_value("Kalkulation!A44")
    return act_axn_k(x_val + a44_val, max(0, n - a44_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d44():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A44_value = get_cell_value("Kalkulation!A44")
    return act_axn_k(x + A44_value, max(0, t - A44_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e44():
    B44 = get_cell_value("Kalkulation!B44")
    P_xt = get_excel_global('P_xt')
    D44 = get_cell_value("Kalkulation!D44")
    gamma2 = get_excel_global('gamma2')
    C44 = get_cell_value("Kalkulation!C44")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B44 - P_xt * D44 + gamma2 * (C44 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D44)

def fkt_kalkulation_f44():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E44")
    return VS * value

def fkt_kalkulation_g44():
    gamma3 = get_excel_global('gamma3')
    b44 = get_cell_value("Kalkulation!B44")
    c44 = get_cell_value("Kalkulation!C44")
    return b44 + gamma3 * c44

def fkt_kalkulation_h44():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F44 = get_cell_value("Kalkulation!F44")
    A44 = get_cell_value("Kalkulation!A44")

    return F44 + alpha * t * BJB * act_axn_k(x + A44, max(5 - A44, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i44():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A44 = get_cell_value("Kalkulation!A44")
    return 1 if (x + A44 >= MinAlterFlex and A44 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j44():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A44 = get_cell_value("Kalkulation!A44")
    I44 = get_cell_value("Kalkulation!I44")
    F44 = get_cell_value("Kalkulation!F44")

    if A44 > n or I44:
        return 0
    else:
        inner_value = VS - F44
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k44():
    h44 = get_cell_value("Kalkulation!H44")
    j44 = get_cell_value("Kalkulation!J44")
    return max(0, h44 - j44)

def fkt_kalkulation_l44():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A44 = get_cell_value("Kalkulation!A44")
    H44 = get_cell_value("Kalkulation!H44")
    G44 = get_cell_value("Kalkulation!G44")
    try:
        if A44 > n:
            return 0
        elif A44 < t:
            return H44 / G44
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b45():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A45 = get_cell_value("Kalkulation!A45")

    if A45 <= n:
        return act_ngr_ax(x + A45, max(0, n - A45), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A45, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c45():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A45_value = get_cell_value("Kalkulation!A45")
    return act_axn_k(x + A45_value, max(0, n - A45_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d45():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A45_value = get_cell_value("Kalkulation!A45")
    return act_axn_k(x + A45_value, max(0, t - A45_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e45():
    B45 = get_cell_value("Kalkulation!B45")
    D45 = get_cell_value("Kalkulation!D45")
    C45 = get_cell_value("Kalkulation!C45")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B45 - P_xt * D45 + gamma2 * (C45 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D45)
    return result

def fkt_kalkulation_f45():
    VS = get_excel_global('VS')
    E45 = get_cell_value("Kalkulation!E45")
    return VS * E45

def fkt_kalkulation_g45():
    gamma3 = get_excel_global('gamma3')
    B45 = get_cell_value("Kalkulation!B45")
    C45 = get_cell_value("Kalkulation!C45")
    return B45 + gamma3 * C45

def fkt_kalkulation_h45():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F45 = get_cell_value("Kalkulation!F45")
    A45 = get_cell_value("Kalkulation!A45")

    return F45 + alpha * t * BJB * act_axn_k(x + A45, max(5 - A45, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i45():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A45 = get_cell_value("Kalkulation!A45")
    return 1 if (x + A45 >= MinAlterFlex and A45 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j45():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A45 = get_cell_value("Kalkulation!A45")
    I45 = get_cell_value("Kalkulation!I45")
    F45 = get_cell_value("Kalkulation!F45")

    if A45 > n or I45:
        return 0
    else:
        result = VS - F45
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k45():
    H45 = get_cell_value("Kalkulation!H45")
    J45 = get_cell_value("Kalkulation!J45")
    return max(0, H45 - J45)

def fkt_kalkulation_l45():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A45 = get_cell_value("Kalkulation!A45")
    H45 = get_cell_value("Kalkulation!H45")
    G45 = get_cell_value("Kalkulation!G45")
    try:
        if A45 > n:
            result = 0
        else:
            if A45 < t:
                result = H45 / G45
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b46():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A46 = get_cell_value("Kalkulation!A46")
    result = act_ngr_ax(x + A46, max(0, n - A46), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A46, Sex, Tafel, Zins) if A46 <= n else 0
    return result

def fkt_kalkulation_c46():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A46 = get_cell_value("Kalkulation!A46")
    return act_axn_k(x + A46, max(0, n - A46), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d46():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A46")
    t_value = t - get_cell_value("Kalkulation!A46")
    max_value = max(0, t_value)

    return act_axn_k(x_value, max_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_e46():
    B46 = get_cell_value("Kalkulation!B46")
    D46 = get_cell_value("Kalkulation!D46")
    C46 = get_cell_value("Kalkulation!C46")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')
    P_xt = get_excel_global('P_xt')
    gamma2 = get_excel_global('gamma2')

    return B46 - P_xt * D46 + gamma2 * (C46 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D46)

def fkt_kalkulation_f46():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E46")
    return VS * value

def fkt_kalkulation_g46():
    gamma3 = get_excel_global('gamma3')
    B46 = get_cell_value("Kalkulation!B46")
    C46 = get_cell_value("Kalkulation!C46")
    return B46 + gamma3 * C46

def fkt_kalkulation_h46():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F46 = get_cell_value("Kalkulation!F46")
    A46 = get_cell_value("Kalkulation!A46")

    return F46 + alpha * t * BJB * act_axn_k(x + A46, max(5 - A46, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i46():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A46 = get_cell_value("Kalkulation!A46")
    return 1 if (x + A46 >= MinAlterFlex and A46 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j46():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A46 = get_cell_value("Kalkulation!A46")
    I46 = get_cell_value("Kalkulation!I46")
    F46 = get_cell_value("Kalkulation!F46")

    if A46 > n or I46:
        return 0
    else:
        inner_value = VS - F46
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k46():
    value = get_cell_value("Kalkulation!H46")
    value2 = get_cell_value("Kalkulation!J46")
    return max(0, value - value2)

def fkt_kalkulation_l46():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A46 = get_cell_value("Kalkulation!A46")
    H46 = get_cell_value("Kalkulation!H46")
    G46 = get_cell_value("Kalkulation!G46")
    try:
        if A46 > n:
            result = 0
        elif A46 < t:
            result = H46 / G46
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b47():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A47 = get_cell_value("Kalkulation!A47")

    if A47 <= n:
        return act_ngr_ax(x + A47, max(0, n - A47), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A47, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c47():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A47_value = get_cell_value("Kalkulation!A47")
    return act_axn_k(x + A47_value, max(0, n - A47_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d47():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A47_value = get_cell_value("Kalkulation!A47")
    return act_axn_k(x + A47_value, max(0, t - A47_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e47():
    B47 = get_cell_value("Kalkulation!B47")
    D47 = get_cell_value("Kalkulation!D47")
    C47 = get_cell_value("Kalkulation!C47")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B47 - P_xt * D47 + gamma2 * (C47 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D47)

def fkt_kalkulation_f47():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E47")
    return VS * value

def fkt_kalkulation_g47():
    gamma3 = get_excel_global('gamma3')
    B47 = get_cell_value("Kalkulation!B47")
    C47 = get_cell_value("Kalkulation!C47")
    return B47 + gamma3 * C47

def fkt_kalkulation_h47():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F47 = get_cell_value("Kalkulation!F47")
    A47 = get_cell_value("Kalkulation!A47")

    return F47 + alpha * t * BJB * act_axn_k(x + A47, max(5 - A47, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i47():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A47 = get_cell_value("Kalkulation!A47")
    return 1 if (x + A47 >= MinAlterFlex and A47 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j47():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A47 = get_cell_value("Kalkulation!A47")
    I47 = get_cell_value("Kalkulation!I47")
    F47 = get_cell_value("Kalkulation!F47")

    if A47 > n or I47:
        return 0
    else:
        inner_value = VS - F47
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k47():
    H47 = get_cell_value("Kalkulation!H47")
    J47 = get_cell_value("Kalkulation!J47")
    return max(0, H47 - J47)

def fkt_kalkulation_l47():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A47 = get_cell_value("Kalkulation!A47")
    H47 = get_cell_value("Kalkulation!H47")
    G47 = get_cell_value("Kalkulation!G47")
    try:
        if A47 > n:
            result = 0
        elif A47 < t:
            result = H47 / G47
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b48():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A48 = get_cell_value("Kalkulation!A48")
    if A48 <= n:
        return act_ngr_ax(x + A48, max(0, n - A48), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A48, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c48():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A48 = get_cell_value("Kalkulation!A48")
    return act_axn_k(x + A48, max(0, n - A48), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d48():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D48")
    A48_val = get_cell_value("Kalkulation!A48")
    t_val = get_cell_value("Kalkulation!A48")

    return act_axn_k(x_val + A48_val, max(0, t_val - A48_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e48():
    gamma2 = get_excel_global('gamma2')
    k = get_excel_global('k')
    n = get_excel_global('n')
    P_xt = get_excel_global('P_xt')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    B48 = get_cell_value("Kalkulation!B48")
    D48 = get_cell_value("Kalkulation!D48")
    C48 = get_cell_value("Kalkulation!C48")

    result = B48 - P_xt * D48 + gamma2 * (C48 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D48)
    return result

def fkt_kalkulation_f48():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E48")
    return VS * value

def fkt_kalkulation_g48():
    gamma3 = get_excel_global('gamma3')
    B48 = get_cell_value("Kalkulation!B48")
    C48 = get_cell_value("Kalkulation!C48")
    return B48 + gamma3 * C48

def fkt_kalkulation_h48():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F48 = get_cell_value("Kalkulation!F48")
    A48 = get_cell_value("Kalkulation!A48")

    return F48 + alpha * t * BJB * act_axn_k(x + A48, max(5 - A48, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i48():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A48 = get_cell_value("Kalkulation!A48")
    return 1 if (x + A48 >= MinAlterFlex and A48 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j48():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A48 = get_cell_value("Kalkulation!A48")
    I48 = get_cell_value("Kalkulation!I48")
    F48 = get_cell_value("Kalkulation!F48")

    if A48 > n or I48:
        return 0
    else:
        inner_value = VS - F48
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k48():
    H48 = get_cell_value("Kalkulation!H48")
    J48 = get_cell_value("Kalkulation!J48")
    return max(0, H48 - J48)

def fkt_kalkulation_l48():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A48 = get_cell_value("Kalkulation!A48")
    H48 = get_cell_value("Kalkulation!H48")
    G48 = get_cell_value("Kalkulation!G48")
    try:
        if A48 > n:
            result = 0
        elif A48 < t:
            result = H48 / G48
        else:
            result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b49():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A49 = get_cell_value("Kalkulation!A49")
    return act_ngr_ax(x + A49, max(0, n - A49), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A49, Sex, Tafel, Zins) if A49 <= n else 0

def fkt_kalkulation_c49():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    A49 = get_cell_value("Kalkulation!A49")
    return act_axn_k(x + A49, max(0, n - A49), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d49():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    A49 = get_cell_value("Kalkulation!A49")
    return act_axn_k(x + A49, max(0, t - A49), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e49():
    B49 = get_cell_value("Kalkulation!B49")
    D49 = get_cell_value("Kalkulation!D49")
    C49 = get_cell_value("Kalkulation!C49")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B49 - P_xt * D49 + gamma2 * (C49 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D49)

def fkt_kalkulation_f49():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E49")
    return VS * value

def fkt_kalkulation_g49():
    gamma3 = get_excel_global('gamma3')
    b49 = get_cell_value("Kalkulation!B49")
    c49 = get_cell_value("Kalkulation!C49")
    return b49 + gamma3 * c49

def fkt_kalkulation_h49():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F49 = get_cell_value("Kalkulation!F49")
    A49 = get_cell_value("Kalkulation!A49")

    return F49 + alpha * t * BJB * act_axn_k(x + A49, max(5 - A49, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i49():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A49 = get_cell_value("Kalkulation!A49")
    return 1 if (x + A49 >= MinAlterFlex and A49 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j49():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A49 = get_cell_value("Kalkulation!A49")
    I49 = get_cell_value("Kalkulation!I49")
    F49 = get_cell_value("Kalkulation!F49")

    if A49 > n or I49:
        return 0
    else:
        inner_value = VS - F49
        result = 1 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k49():
    H49 = get_cell_value("Kalkulation!H49")
    J49 = get_cell_value("Kalkulation!J49")
    return max(0, H49 - J49)

def fkt_kalkulation_l49():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A49 = get_cell_value("Kalkulation!A49")
    H49 = get_cell_value("Kalkulation!H49")
    G49 = get_cell_value("Kalkulation!G49")
    try:
        if A49 > n:
            result = 0
        else:
            if A49 < t:
                result = H49 / G49
            else:
                result = VS
    except:
        result = 0
    return result

def fkt_kalkulation_b50():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A50 = get_cell_value("Kalkulation!A50")

    if A50 <= n:
        return act_ngr_ax(x + A50, max(0, n - A50), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A50, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c50():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A50_value = get_cell_value("Kalkulation!A50")
    return act_axn_k(x + A50_value, max(0, n - A50_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d50():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D50")
    a50_val = get_cell_value("Kalkulation!A50")
    t_val = get_cell_value("Kalkulation!A50")

    return act_axn_k(x_val + a50_val, max(0, t_val - a50_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e50():
    B50 = get_cell_value("Kalkulation!B50")
    D50 = get_cell_value("Kalkulation!D50")
    C50 = get_cell_value("Kalkulation!C50")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B50 - P_xt * D50 + gamma2 * (C50 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D50)

def fkt_kalkulation_f50():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E50")
    return VS * value

def fkt_kalkulation_g50():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B50") + gamma3 * get_cell_value("Kalkulation!C50")
    return value

def fkt_kalkulation_h50():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F50 = get_cell_value("Kalkulation!F50")
    A50 = get_cell_value("Kalkulation!A50")

    return F50 + alpha * t * BJB * act_axn_k(x + A50, max(5 - A50, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i50():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A50_value = get_cell_value("Kalkulation!A50")
    return 1 if (x + A50_value >= MinAlterFlex and A50_value >= n - MinRLZFlex) else 0

def fkt_kalkulation_j50():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A50 = get_cell_value("Kalkulation!A50")
    I50 = get_cell_value("Kalkulation!I50")
    F50 = get_cell_value("Kalkulation!F50")
    return 0 if (A50 > n) or I50 else min(150, max(50, 0.01 * (VS - F50)))

def fkt_kalkulation_k50():
    h50 = get_cell_value("Kalkulation!H50")
    j50 = get_cell_value("Kalkulation!J50")
    return max(0, h50 - j50)

def fkt_kalkulation_l50():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A50 = get_cell_value("Kalkulation!A50")
    H50 = get_cell_value("Kalkulation!H50")
    G50 = get_cell_value("Kalkulation!G50")

    try:
        if A50 > n:
            return 0
        elif A50 < t:
            return H50 / G50
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b51():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A51 = get_cell_value("Kalkulation!A51")

    if A51 <= n:
        return act_ngr_ax(x + A51, max(0, n - A51), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A51, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c51():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A51_value = get_cell_value("Kalkulation!A51")
    return act_axn_k(x + A51_value, max(0, n - A51_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d51():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A51_value = get_cell_value("Kalkulation!A51")
    return act_axn_k(x + A51_value, max(0, t - A51_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e51():
    B51 = get_cell_value("Kalkulation!B51")
    P_xt = get_excel_global('P_xt')
    D51 = get_cell_value("Kalkulation!D51")
    gamma2 = get_excel_global('gamma2')
    C51 = get_cell_value("Kalkulation!C51")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B51 - P_xt * D51 + gamma2 * (C51 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D51)

def fkt_kalkulation_f51():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E51")
    return VS * value

def fkt_kalkulation_g51():
    gamma3 = get_excel_global('gamma3')
    return get_cell_value("Kalkulation!B51") + gamma3 * get_cell_value("Kalkulation!C51")

def fkt_kalkulation_h51():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F51 = get_cell_value("Kalkulation!F51")
    A51 = get_cell_value("Kalkulation!A51")

    return F51 + alpha * t * BJB * act_axn_k(x + A51, max(5 - A51, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i51():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A51 = get_cell_value("Kalkulation!A51")
    return 1 if (x + A51 >= MinAlterFlex and A51 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j51():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A51 = get_cell_value("Kalkulation!A51")
    I51 = get_cell_value("Kalkulation!I51")
    F51 = get_cell_value("Kalkulation!F51")

    if A51 > n or I51:
        return 0
    else:
        inner_value = VS - F51
        result = inner_value * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k51():
    h51 = get_cell_value("Kalkulation!H51")
    j51 = get_cell_value("Kalkulation!J51")
    return max(0, h51 - j51)

def fkt_kalkulation_l51():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A51 = get_cell_value("Kalkulation!A51")
    H51 = get_cell_value("Kalkulation!H51")
    G51 = get_cell_value("Kalkulation!G51")
    try:
        if A51 > n:
            return 0
        elif A51 < t:
            return H51 / G51
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b52():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A52 = get_cell_value("Kalkulation!A52")

    if A52 <= n:
        return act_ngr_ax(x + A52, max(0, n - A52), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A52, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c52():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A52 = get_cell_value("Kalkulation!A52")
    return act_axn_k(x + A52, max(0, n - A52), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d52():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A52_value = get_cell_value("Kalkulation!A52")
    return act_axn_k(x + A52_value, max(0, t - A52_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e52():
    B52 = get_cell_value("Kalkulation!B52")
    D52 = get_cell_value("Kalkulation!D52")
    C52 = get_cell_value("Kalkulation!C52")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B52 - P_xt * D52 + gamma2 * (C52 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D52)

def fkt_kalkulation_f52():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E52")
    return VS * value

def fkt_kalkulation_g52():
    gamma3 = get_excel_global('gamma3')
    B52 = get_cell_value("Kalkulation!B52")
    C52 = get_cell_value("Kalkulation!C52")
    return B52 + gamma3 * C52

def fkt_kalkulation_h52():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F52 = get_cell_value("Kalkulation!F52")
    A52 = get_cell_value("Kalkulation!A52")

    return F52 + alpha * t * BJB * act_axn_k(x + A52, max(5 - A52, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i52():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A52 = get_cell_value("Kalkulation!A52")
    return 1 if (x + A52 >= MinAlterFlex and A52 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j52():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A52 = get_cell_value("Kalkulation!A52")
    I52 = get_cell_value("Kalkulation!I52")
    F52 = get_cell_value("Kalkulation!F52")
    if A52 > n or I52:
        return 0
    else:
        inner_value = VS - F52
        result = 0.01 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k52():
    h52 = get_cell_value("Kalkulation!H52")
    j52 = get_cell_value("Kalkulation!J52")
    return max(0, h52 - j52)

def fkt_kalkulation_l52():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A52 = get_cell_value("Kalkulation!A52")
    H52 = get_cell_value("Kalkulation!H52")
    G52 = get_cell_value("Kalkulation!G52")
    try:
        if A52 > n:
            return 0
        elif A52 < t:
            return H52 / G52
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b53():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A53 = get_cell_value("Kalkulation!A53")

    if A53 <= n:
        return act_ngr_ax(x + A53, max(0, n - A53), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A53, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c53():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A53_value = get_cell_value("Kalkulation!A53")
    return act_axn_k(x + A53_value, max(0, n - A53_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d53():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A53 = get_cell_value("Kalkulation!A53")
    return act_axn_k(x + A53, max(0, t - A53), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e53():
    B53 = get_cell_value("Kalkulation!B53")
    D53 = get_cell_value("Kalkulation!D53")
    C53 = get_cell_value("Kalkulation!C53")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B53 - P_xt * D53 + gamma2 * (C53 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D53)
    return result

def fkt_kalkulation_f53():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E53")
    return VS * value

def fkt_kalkulation_g53():
    gamma3 = get_excel_global('gamma3')
    B53 = get_cell_value("Kalkulation!B53")
    C53 = get_cell_value("Kalkulation!C53")
    return B53 + gamma3 * C53

def fkt_kalkulation_h53():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F53 = get_cell_value("Kalkulation!F53")
    A53 = get_cell_value("Kalkulation!A53")

    result = F53 + alpha * t * BJB * act_axn_k(x + A53, max(5 - A53, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)
    return result

def fkt_kalkulation_i53():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A53 = get_cell_value("Kalkulation!A53")
    return 1 if (x + A53 >= MinAlterFlex and A53 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j53():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A53 = get_cell_value("Kalkulation!A53")
    F53 = get_cell_value("Kalkulation!F53")
    I53 = get_cell_value("Kalkulation!I53")
    if A53 > n or I53:
        return 0
    else:
        inner_value = VS - F53
        result = 0.01 * inner_value
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k53():
    H53 = get_cell_value("Kalkulation!H53")
    J53 = get_cell_value("Kalkulation!J53")
    return max(0, H53 - J53)

def fkt_kalkulation_l53():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A53 = get_cell_value("Kalkulation!A53")
    H53 = get_cell_value("Kalkulation!H53")
    G53 = get_cell_value("Kalkulation!G53")
    try:
        if A53 > n:
            return 0
        elif A53 < t:
            return H53 / G53
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b54():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A54 = get_cell_value("Kalkulation!A54")

    if A54 <= n:
        return act_ngr_ax(x + A54, max(0, n - A54), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A54, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c54():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A54_value = get_cell_value("Kalkulation!A54")
    return act_axn_k(x + A54_value, max(0, n - A54_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d54():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A54")
    t_value = t - get_cell_value("Kalkulation!A54")
    max_value = max(0, t_value)

    return act_axn_k(alter=x_value, n=max_value, sex=Sex, tafel=Tafel, zins=Zins, k=1)

def fkt_kalkulation_e54():
    B54 = get_cell_value("Kalkulation!B54")
    D54 = get_cell_value("Kalkulation!D54")
    C54 = get_cell_value("Kalkulation!C54")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B54 - P_xt * D54 + gamma2 * (C54 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D54)

def fkt_kalkulation_f54():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E54")
    return VS * value

def fkt_kalkulation_g54():
    gamma3 = get_excel_global('gamma3')
    B54 = get_cell_value("Kalkulation!B54")
    C54 = get_cell_value("Kalkulation!C54")
    return B54 + gamma3 * C54

def fkt_kalkulation_h54():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F54 = get_cell_value("Kalkulation!F54")
    A54 = get_cell_value("Kalkulation!A54")

    return F54 + alpha * t * BJB * act_axn_k(x + A54, max(5 - A54, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i54():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A54 = get_cell_value("Kalkulation!A54")
    return 1 if (x + A54 >= MinAlterFlex and A54 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j54():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A54 = get_cell_value("Kalkulation!A54")
    I54 = get_cell_value("Kalkulation!I54")
    F54 = get_cell_value("Kalkulation!F54")

    if A54 > n or I54:
        return 0
    else:
        inner_value = VS - F54
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k54():
    h54 = get_cell_value("Kalkulation!H54")
    j54 = get_cell_value("Kalkulation!J54")
    return max(0, h54 - j54)

def fkt_kalkulation_l54():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A54 = get_cell_value("Kalkulation!A54")
    H54 = get_cell_value("Kalkulation!H54")
    G54 = get_cell_value("Kalkulation!G54")
    try:
        if A54 > n:
            return 0
        elif A54 < t:
            return H54 / G54
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b55():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A55 = get_cell_value("Kalkulation!A55")

    if A55 <= n:
        return act_ngr_ax(x + A55, max(0, n - A55), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A55, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c55():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A55_value = get_cell_value("Kalkulation!A55")
    return act_axn_k(x + A55_value, max(0, n - A55_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d55():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A55_value = get_cell_value("Kalkulation!A55")
    return act_axn_k(x + A55_value, max(0, t - A55_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e55():
    B55 = get_cell_value("Kalkulation!B55")
    D55 = get_cell_value("Kalkulation!D55")
    C55 = get_cell_value("Kalkulation!C55")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B55 - P_xt * D55 + gamma2 * (C55 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D55)

def fkt_kalkulation_f55():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E55")
    return VS * value

def fkt_kalkulation_g55():
    gamma3 = get_excel_global('gamma3')
    B55 = get_cell_value("Kalkulation!B55")
    C55 = get_cell_value("Kalkulation!C55")
    return B55 + gamma3 * C55

def fkt_kalkulation_h55():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F55 = get_cell_value("Kalkulation!F55")
    A55 = get_cell_value("Kalkulation!A55")

    return F55 + alpha * t * BJB * act_axn_k(x + A55, max(5 - A55, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i55():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A55 = get_cell_value("Kalkulation!A55")
    return 1 if (x + A55 >= MinAlterFlex and A55 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j55():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A55 = get_cell_value("Kalkulation!A55")
    I55 = get_cell_value("Kalkulation!I55")
    F55 = get_cell_value("Kalkulation!F55")
    return 0 if (A55 > n) or I55 else min(150, max(50, 0.01 * (VS - F55)))

def fkt_kalkulation_k55():
    h55 = get_cell_value("Kalkulation!H55")
    j55 = get_cell_value("Kalkulation!J55")
    return max(0, h55 - j55)

def fkt_kalkulation_l55():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A55 = get_cell_value("Kalkulation!A55")
    H55 = get_cell_value("Kalkulation!H55")
    G55 = get_cell_value("Kalkulation!G55")

    try:
        if A55 > n:
            result = 0
        elif A55 < t:
            result = H55 / G55
        else:
            result = VS
    except:
        result = 0

    return result

def fkt_kalkulation_b56():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A56 = get_cell_value("Kalkulation!A56")

    if A56 <= n:
        return act_ngr_ax(x + A56, max(0, n - A56), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A56, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c56():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A56_value = get_cell_value("Kalkulation!A56")
    return act_axn_k(x + A56_value, max(0, n - A56_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d56():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A56 = get_cell_value("Kalkulation!A56")
    return act_axn_k(x + A56, max(0, t - A56), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e56() -> float:
    B56 = get_cell_value("Kalkulation!B56")
    D56 = get_cell_value("Kalkulation!D56")
    C56 = get_cell_value("Kalkulation!C56")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B56 - P_xt * D56 + gamma2 * (C56 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D56)

def fkt_kalkulation_f56():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E56")
    return VS * value

def fkt_kalkulation_g56():
    gamma3 = get_excel_global('gamma3')
    return get_cell_value("Kalkulation!B56") + gamma3 * get_cell_value("Kalkulation!C56")

def fkt_kalkulation_h56():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F56 = get_cell_value("Kalkulation!F56")
    A56 = get_cell_value("Kalkulation!A56")

    return F56 + alpha * t * BJB * act_axn_k(x + A56, max(5 - A56, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i56():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A56 = get_cell_value("Kalkulation!A56")
    return 1 if (x + A56 >= MinAlterFlex and A56 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j56():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A56 = get_cell_value("Kalkulation!A56")
    I56 = get_cell_value("Kalkulation!I56")
    F56 = get_cell_value("Kalkulation!F56")

    if A56 > n or I56:
        return 0
    else:
        inner_value = VS - F56
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k56():
    h56 = get_cell_value("Kalkulation!H56")
    j56 = get_cell_value("Kalkulation!J56")
    return max(0, h56 - j56)

def fkt_kalkulation_l56():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A56 = get_cell_value("Kalkulation!A56")
    H56 = get_cell_value("Kalkulation!H56")
    G56 = get_cell_value("Kalkulation!G56")
    try:
        if A56 > n:
            return 0
        elif A56 < t:
            return H56 / G56
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b57():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A57 = get_cell_value("Kalkulation!A57")

    if A57 <= n:
        return act_ngr_ax(x + A57, max(0, n - A57), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A57, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c57():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A57 = get_cell_value("Kalkulation!A57")
    return act_axn_k(x + A57, max(0, n - A57), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d57():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A57_value = get_cell_value("Kalkulation!A57")
    return act_axn_k(x + A57_value, max(0, t - A57_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e57():
    B57 = get_cell_value("Kalkulation!B57")
    D57 = get_cell_value("Kalkulation!D57")
    C57 = get_cell_value("Kalkulation!C57")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B57 - P_xt * D57 + gamma2 * (C57 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D57)
    return result

def fkt_kalkulation_f57():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E57")
    return VS * value

def fkt_kalkulation_g57():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B57") + gamma3 * get_cell_value("Kalkulation!C57")
    return value

def fkt_kalkulation_h57() -> float:
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F57 = get_cell_value("Kalkulation!F57")
    A57 = get_cell_value("Kalkulation!A57")

    return F57 + alpha * t * BJB * act_axn_k(x + A57, max(5 - A57, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i57():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A57 = get_cell_value("Kalkulation!A57")
    return 1 if (x + A57 >= MinAlterFlex and A57 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j57():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A57 = get_cell_value("Kalkulation!A57")
    I57 = get_cell_value("Kalkulation!I57")
    F57 = get_cell_value("Kalkulation!F57")
    if A57 > n or I57:
        return 0
    else:
        inner_value = VS - F57
        result = inner_value * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k57():
    value = get_cell_value("Kalkulation!H57")
    value2 = get_cell_value("Kalkulation!J57")
    return max(0, value - value2)

def fkt_kalkulation_l57():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A57 = get_cell_value("Kalkulation!A57")
    H57 = get_cell_value("Kalkulation!H57")
    G57 = get_cell_value("Kalkulation!G57")
    try:
        if A57 > n:
            return 0
        elif A57 < t:
            return H57 / G57
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b58():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A58 = get_cell_value("Kalkulation!A58")

    if A58 <= n:
        return act_ngr_ax(x + A58, max(0, n - A58), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A58, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c58():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A58_value = get_cell_value("Kalkulation!A58")
    return act_axn_k(x + A58_value, max(0, n - A58_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d58():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A58_value = get_cell_value("Kalkulation!A58")
    return act_axn_k(x + A58_value, max(0, t - A58_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e58():
    B58 = get_cell_value("Kalkulation!B58")
    D58 = get_cell_value("Kalkulation!D58")
    C58 = get_cell_value("Kalkulation!C58")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B58 - P_xt * D58 + gamma2 * (C58 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D58)

def fkt_kalkulation_f58():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E58")
    return VS * value

def fkt_kalkulation_g58():
    gamma3 = get_excel_global('gamma3')
    B58 = get_cell_value("Kalkulation!B58")
    C58 = get_cell_value("Kalkulation!C58")
    return B58 + gamma3 * C58

def fkt_kalkulation_h58():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F58 = get_cell_value("Kalkulation!F58")
    A58 = get_cell_value("Kalkulation!A58")

    return F58 + alpha * t * BJB * act_axn_k(x + A58, max(5 - A58, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i58():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A58 = get_cell_value("Kalkulation!A58")
    return 1 if (x + A58 >= MinAlterFlex and A58 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j58():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A58 = get_cell_value("Kalkulation!A58")
    I58 = get_cell_value("Kalkulation!I58")
    F58 = get_cell_value("Kalkulation!F58")

    if A58 > n or I58:
        return 0
    else:
        inner_value = VS - F58
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k58():
    H58 = get_cell_value("Kalkulation!H58")
    J58 = get_cell_value("Kalkulation!J58")
    return max(0, H58 - J58)

def fkt_kalkulation_l58():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A58 = get_cell_value("Kalkulation!A58")
    H58 = get_cell_value("Kalkulation!H58")
    G58 = get_cell_value("Kalkulation!G58")
    try:
        if A58 > n:
            return 0
        elif A58 < t:
            return H58 / G58
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b59():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A59 = get_cell_value("Kalkulation!A59")

    if A59 <= n:
        return act_ngr_ax(x + A59, max(0, n - A59), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A59, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c59():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A59_value = get_cell_value("Kalkulation!A59")
    return act_axn_k(x + A59_value, max(0, n - A59_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d59():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A59_value = get_cell_value("Kalkulation!A59")
    return act_axn_k(x + A59_value, max(0, t - A59_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e59():
    B59 = get_cell_value("Kalkulation!B59")
    D59 = get_cell_value("Kalkulation!D59")
    C59 = get_cell_value("Kalkulation!C59")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B59 - P_xt * D59 + gamma2 * (C59 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D59)
    return result

def fkt_kalkulation_f59():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E59")
    return VS * value

def fkt_kalkulation_g59():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B59") + gamma3 * get_cell_value("Kalkulation!C59")
    return value

def fkt_kalkulation_h59():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F59 = get_cell_value("Kalkulation!F59")
    A59 = get_cell_value("Kalkulation!A59")

    return F59 + alpha * t * BJB * act_axn_k(x + A59, max(5 - A59, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i59():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A59 = get_cell_value("Kalkulation!A59")
    return 1 if (x + A59 >= MinAlterFlex and A59 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j59():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A59 = get_cell_value("Kalkulation!A59")
    I59 = get_cell_value("Kalkulation!I59")
    F59 = get_cell_value("Kalkulation!F59")
    if A59 > n or I59:
        return 0
    else:
        return min(150, max(50, 0.01 * (VS - F59)))

def fkt_kalkulation_k59():
    h59 = get_cell_value("Kalkulation!H59")
    j59 = get_cell_value("Kalkulation!J59")
    return max(0, h59 - j59)

def fkt_kalkulation_l59():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A59 = get_cell_value("Kalkulation!A59")
    H59 = get_cell_value("Kalkulation!H59")
    G59 = get_cell_value("Kalkulation!G59")
    try:
        if A59 > n:
            return 0
        elif A59 < t:
            return H59 / G59
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b60():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A60 = get_cell_value("Kalkulation!A60")

    if A60 <= n:
        return act_ngr_ax(x + A60, max(0, n - A60), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A60, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c60():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A60_value = get_cell_value("Kalkulation!A60")
    alter = x + A60_value
    n_value = max(0, n - A60_value)

    return act_axn_k(alter, n_value, Sex, Tafel, Zins, 1)

def fkt_kalkulation_d60():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A60_value = get_cell_value("Kalkulation!A60")
    return act_axn_k(x + A60_value, max(0, t - A60_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e60():
    B60 = get_cell_value("Kalkulation!B60")
    D60 = get_cell_value("Kalkulation!D60")
    C60 = get_cell_value("Kalkulation!C60")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')
    P_xt = get_excel_global('P_xt')
    gamma2 = get_excel_global('gamma2')

    return B60 - P_xt * D60 + gamma2 * (C60 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D60)

def fkt_kalkulation_f60():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E60")
    return VS * value

def fkt_kalkulation_g60():
    gamma3 = get_excel_global('gamma3')
    B60 = get_cell_value("Kalkulation!B60")
    C60 = get_cell_value("Kalkulation!C60")
    return B60 + gamma3 * C60

def fkt_kalkulation_h60():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F60 = get_cell_value("Kalkulation!F60")
    A60 = get_cell_value("Kalkulation!A60")

    return F60 + alpha * t * BJB * act_axn_k(x + A60, max(5 - A60, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i60():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A60 = get_cell_value("Kalkulation!A60")
    return 1 if (x + A60 >= MinAlterFlex and A60 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j60():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A60 = get_cell_value("Kalkulation!A60")
    I60 = get_cell_value("Kalkulation!I60")
    F60 = get_cell_value("Kalkulation!F60")
    return 0 if (A60 > n) or I60 else min(150, max(50, 0.01 * (VS - F60)))

def fkt_kalkulation_k60():
    h60 = get_cell_value("Kalkulation!H60")
    j60 = get_cell_value("Kalkulation!J60")
    return max(0, h60 - j60)

def fkt_kalkulation_l60():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A60 = get_cell_value("Kalkulation!A60")
    H60 = get_cell_value("Kalkulation!H60")
    G60 = get_cell_value("Kalkulation!G60")
    try:
        if A60 > n:
            return 0
        elif A60 < t:
            return H60 / G60
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b61():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A61 = get_cell_value("Kalkulation!A61")

    if A61 <= n:
        return act_ngr_ax(x + A61, max(0, n - A61), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A61, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c61():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!C61")
    a61_val = get_cell_value("Kalkulation!A61")
    return act_axn_k(x_val + a61_val, max(0, n - a61_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d61():
    x = get_excel_global('x')
    t = get_excel_global('t')
    A61 = get_cell_value("Kalkulation!A61")
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    return act_axn_k(x + A61, max(0, t - A61), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e61():
    B61 = get_cell_value("Kalkulation!B61")
    P_xt = get_excel_global('P_xt')
    D61 = get_cell_value("Kalkulation!D61")
    gamma2 = get_excel_global('gamma2')
    C61 = get_cell_value("Kalkulation!C61")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B61 - P_xt * D61 + gamma2 * (C61 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D61)

def fkt_kalkulation_f61():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E61")
    return VS * value

def fkt_kalkulation_g61():
    gamma3 = get_excel_global('gamma3')
    B61 = get_cell_value("Kalkulation!B61")
    C61 = get_cell_value("Kalkulation!C61")
    return B61 + gamma3 * C61

def fkt_kalkulation_h61():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F61 = get_cell_value("Kalkulation!F61")
    A61 = get_cell_value("Kalkulation!A61")

    return F61 + alpha * t * BJB * act_axn_k(x + A61, max(5 - A61, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i61():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A61 = get_cell_value("Kalkulation!A61")
    return 1 if (x + A61 >= MinAlterFlex and A61 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j61():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A61 = get_cell_value("Kalkulation!A61")
    I61 = get_cell_value("Kalkulation!I61")
    F61 = get_cell_value("Kalkulation!F61")
    return 0 if (A61 > n) or I61 else min(150, max(50, 0.01 * (VS - F61)))

def fkt_kalkulation_k61():
    h61 = get_cell_value("Kalkulation!H61")
    j61 = get_cell_value("Kalkulation!J61")
    return max(0, h61 - j61)

def fkt_kalkulation_l61():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A61 = get_cell_value("Kalkulation!A61")
    H61 = get_cell_value("Kalkulation!H61")
    G61 = get_cell_value("Kalkulation!G61")
    try:
        if A61 > n:
            return 0
        elif A61 < t:
            return H61 / G61
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b62():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A62 = get_cell_value("Kalkulation!A62")

    if A62 <= n:
        return act_ngr_ax(x + A62, max(0, n - A62), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A62, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c62():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A62_value = get_cell_value("Kalkulation!A62")
    return act_axn_k(x + A62_value, max(0, n - A62_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d62():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_val = get_cell_value("Kalkulation!D62")
    a62_val = get_cell_value("Kalkulation!A62")
    return act_axn_k(x_val + a62_val, max(0, t - a62_val), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e62():
    B62 = get_cell_value("Kalkulation!B62")
    P_xt = get_excel_global('P_xt')
    D62 = get_cell_value("Kalkulation!D62")
    gamma2 = get_excel_global('gamma2')
    C62 = get_cell_value("Kalkulation!C62")
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B62 - P_xt * D62 + gamma2 * (C62 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D62)

def fkt_kalkulation_f62():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E62")
    return VS * value

def fkt_kalkulation_g62():
    gamma3 = get_excel_global('gamma3')
    B62 = get_cell_value("Kalkulation!B62")
    C62 = get_cell_value("Kalkulation!C62")
    return B62 + gamma3 * C62

def fkt_kalkulation_h62():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F62 = get_cell_value("Kalkulation!F62")
    A62 = get_cell_value("Kalkulation!A62")

    return F62 + alpha * t * BJB * act_axn_k(x + A62, max(5 - A62, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i62():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A62 = get_cell_value("Kalkulation!A62")
    return 1 if (x + A62 >= MinAlterFlex and A62 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j62():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A62 = get_cell_value("Kalkulation!A62")
    F62 = get_cell_value("Kalkulation!F62")
    I62 = get_cell_value("Kalkulation!I62")
    if A62 > n or I62:
        return 0
    else:
        inner_value = VS - F62
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k62():
    H62 = get_cell_value("Kalkulation!H62")
    J62 = get_cell_value("Kalkulation!J62")
    return max(0, H62 - J62)

def fkt_kalkulation_l62():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A62 = get_cell_value("Kalkulation!A62")
    H62 = get_cell_value("Kalkulation!H62")
    G62 = get_cell_value("Kalkulation!G62")
    try:
        if A62 > n:
            return 0
        elif A62 < t:
            return H62 / G62
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b63():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')
    A63 = get_cell_value("Kalkulation!A63")
    if A63 <= n:
        return act_ngr_ax(x + A63, max(0, n - A63), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A63, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c63():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A63_value = get_cell_value("Kalkulation!A63")
    return act_axn_k(x + A63_value, max(0, n - A63_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d63():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A63_value = get_cell_value("Kalkulation!A63")
    return act_axn_k(x + A63_value, max(0, t - A63_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e63():
    B63 = get_cell_value("Kalkulation!B63")
    D63 = get_cell_value("Kalkulation!D63")
    C63 = get_cell_value("Kalkulation!C63")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    result = B63 - P_xt * D63 + gamma2 * (C63 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D63)
    return result

def fkt_kalkulation_f63():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E63")
    return VS * value

def fkt_kalkulation_g63():
    gamma3 = get_excel_global('gamma3')
    B63 = get_cell_value("Kalkulation!B63")
    C63 = get_cell_value("Kalkulation!C63")
    return B63 + gamma3 * C63

def fkt_kalkulation_h63():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F63 = get_cell_value("Kalkulation!F63")
    A63 = get_cell_value("Kalkulation!A63")

    return F63 + alpha * t * BJB * act_axn_k(x + A63, max(5 - A63, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i63():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A63 = get_cell_value("Kalkulation!A63")
    return 1 if (x + A63 >= MinAlterFlex and A63 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j63():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A63 = get_cell_value("Kalkulation!A63")
    F63 = get_cell_value("Kalkulation!F63")
    I63 = get_cell_value("Kalkulation!I63")
    if A63 > n or I63:
        return 0
    else:
        inner_value = VS - F63
        result = min(150, max(50, 0.01 * inner_value))
        return result

def fkt_kalkulation_k63():
    H63 = get_cell_value("Kalkulation!H63")
    J63 = get_cell_value("Kalkulation!J63")
    return max(0, H63 - J63)

def fkt_kalkulation_l63():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A63 = get_cell_value("Kalkulation!A63")
    H63 = get_cell_value("Kalkulation!H63")
    G63 = get_cell_value("Kalkulation!G63")
    try:
        if A63 > n:
            return 0
        elif A63 < t:
            return H63 / G63
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b64():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A64 = get_cell_value("Kalkulation!A64")

    if A64 <= n:
        return act_ngr_ax(x + A64, max(0, n - A64), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A64, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c64():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A64_value = get_cell_value("Kalkulation!A64")
    return act_axn_k(x + A64_value, max(0, n - A64_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d64():
    x = get_excel_global('x')
    t = get_excel_global('t')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A64_value = get_cell_value("Kalkulation!A64")
    return act_axn_k(x + A64_value, max(0, t - A64_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e64():
    B64 = get_cell_value("Kalkulation!B64")
    D64 = get_cell_value("Kalkulation!D64")
    C64 = get_cell_value("Kalkulation!C64")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B64 - P_xt * D64 + gamma2 * (C64 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D64)

def fkt_kalkulation_f64():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E64")
    return VS * value

def fkt_kalkulation_g64():
    gamma3 = get_excel_global('gamma3')
    B64 = get_cell_value("Kalkulation!B64")
    C64 = get_cell_value("Kalkulation!C64")
    return B64 + gamma3 * C64

def fkt_kalkulation_h64():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F64 = get_cell_value("Kalkulation!F64")
    A64 = get_cell_value("Kalkulation!A64")

    result = F64 + alpha * t * BJB * act_axn_k(x + A64, max(5 - A64, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)
    return result

def fkt_kalkulation_i64():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A64 = get_cell_value("Kalkulation!A64")
    return 1 if (x + A64 >= MinAlterFlex and A64 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j64():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A64 = get_cell_value("Kalkulation!A64")
    I64 = get_cell_value("Kalkulation!I64")
    F64 = get_cell_value("Kalkulation!F64")
    return 0 if (A64 > n) or I64 else min(150, max(50, 0.01 * (VS - F64)))

def fkt_kalkulation_k64():
    H64 = get_cell_value("Kalkulation!H64")
    J64 = get_cell_value("Kalkulation!J64")
    return max(0, H64 - J64)

def fkt_kalkulation_l64():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A64 = get_cell_value("Kalkulation!A64")
    H64 = get_cell_value("Kalkulation!H64")
    G64 = get_cell_value("Kalkulation!G64")
    try:
        if A64 > n:
            return 0
        elif A64 < t:
            return H64 / G64
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b65():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A65 = get_cell_value("Kalkulation!A65")

    if A65 <= n:
        return act_ngr_ax(x + A65, max(0, n - A65), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A65, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c65():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A65_value = get_cell_value("Kalkulation!A65")
    return act_axn_k(x + A65_value, max(0, n - A65_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d65():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    x_value = x + get_cell_value("Kalkulation!A65")
    t_value = t - get_cell_value("Kalkulation!A65")
    max_value = max(0, t_value)

    return act_axn_k(alter=x_value, n=max_value, sex=Sex, tafel=Tafel, zins=Zins, k=1)

def fkt_kalkulation_e65():
    B65 = get_cell_value("Kalkulation!B65")
    D65 = get_cell_value("Kalkulation!D65")
    C65 = get_cell_value("Kalkulation!C65")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B65 - P_xt * D65 + gamma2 * (C65 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D65)

def fkt_kalkulation_f65():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E65")
    return VS * value

def fkt_kalkulation_g65():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B65") + gamma3 * get_cell_value("Kalkulation!C65")
    return value

def fkt_kalkulation_h65():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F65 = get_cell_value("Kalkulation!F65")
    A65 = get_cell_value("Kalkulation!A65")

    return F65 + alpha * t * BJB * act_axn_k(x + A65, max(5 - A65, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i65():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A65 = get_cell_value("Kalkulation!A65")
    return 1 if (x + A65 >= MinAlterFlex and A65 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j65():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A65 = get_cell_value("Kalkulation!A65")
    I65 = get_cell_value("Kalkulation!I65")
    F65 = get_cell_value("Kalkulation!F65")
    return 0 if (A65 > n) or I65 else min(150, max(50, 0.01 * (VS - F65)))

def fkt_kalkulation_k65():
    h65 = get_cell_value("Kalkulation!H65")
    j65 = get_cell_value("Kalkulation!J65")
    return max(0, h65 - j65)

def fkt_kalkulation_l65():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A65 = get_cell_value("Kalkulation!A65")
    H65 = get_cell_value("Kalkulation!H65")
    G65 = get_cell_value("Kalkulation!G65")
    try:
        if A65 > n:
            return 0
        elif A65 < t:
            return H65 / G65
        else:
            return VS
    except:
        return 0

def fkt_kalkulation_b66():
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A66 = get_cell_value("Kalkulation!A66")

    if A66 <= n:
        return act_ngr_ax(x + A66, max(0, n - A66), Sex, Tafel, Zins) + act_dx(x + n, Sex, Tafel, Zins) / act_dx(x + A66, Sex, Tafel, Zins)
    else:
        return 0

def fkt_kalkulation_c66():
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')

    A66_value = get_cell_value("Kalkulation!A66")
    return act_axn_k(x + A66_value, max(0, n - A66_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_d66():
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    A66_value = get_cell_value("Kalkulation!A66")
    return act_axn_k(x + A66_value, max(0, t - A66_value), Sex, Tafel, Zins, 1)

def fkt_kalkulation_e66():
    B66 = get_cell_value("Kalkulation!B66")
    D66 = get_cell_value("Kalkulation!D66")
    C66 = get_cell_value("Kalkulation!C66")
    gamma2 = get_excel_global('gamma2')
    P_xt = get_excel_global('P_xt')
    x = get_excel_global('x')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    Tafel = get_excel_global('Tafel')
    Zins = get_excel_global('Zins')
    t = get_excel_global('t')

    return B66 - P_xt * D66 + gamma2 * (C66 - act_axn_k(x, n, Sex, Tafel, Zins, 1) / act_axn_k(x, t, Sex, Tafel, Zins, 1) * D66)

def fkt_kalkulation_f66():
    VS = get_excel_global('VS')
    value = get_cell_value("Kalkulation!E66")
    return VS * value

def fkt_kalkulation_g66():
    gamma3 = get_excel_global('gamma3')
    value = get_cell_value("Kalkulation!B66") + gamma3 * get_cell_value("Kalkulation!C66")
    return value

def fkt_kalkulation_h66():
    alpha = get_excel_global('alpha')
    BJB = get_excel_global('BJB')
    k = get_excel_global('k')
    n = get_excel_global('n')
    Sex = get_excel_global('Sex')
    t = get_excel_global('t')
    Tafel = get_excel_global('Tafel')
    x = get_excel_global('x')
    Zins = get_excel_global('Zins')

    F66 = get_cell_value("Kalkulation!F66")
    A66 = get_cell_value("Kalkulation!A66")

    return F66 + alpha * t * BJB * act_axn_k(x + A66, max(5 - A66, 0), Sex, Tafel, Zins, 1) / act_axn_k(x, 5, Sex, Tafel, Zins, 1)

def fkt_kalkulation_i66():
    MinAlterFlex = get_excel_global('MinAlterFlex')
    MinRLZFlex = get_excel_global('MinRLZFlex')
    n = get_excel_global('n')
    t = get_excel_global('t')
    x = get_excel_global('x')
    A66 = get_cell_value("Kalkulation!A66")
    return 1 if (x + A66 >= MinAlterFlex and A66 >= n - MinRLZFlex) else 0

def fkt_kalkulation_j66():
    n = get_excel_global('n')
    VS = get_excel_global('VS')
    A66 = get_cell_value("Kalkulation!A66")
    I66 = get_cell_value("Kalkulation!I66")
    F66 = get_cell_value("Kalkulation!F66")

    if A66 > n or I66:
        return 0
    else:
        result = VS - F66
        result = result * 0.01
        result = max(50, result)
        result = min(150, result)
        return result

def fkt_kalkulation_k66():
    h66 = get_cell_value("Kalkulation!H66")
    j66 = get_cell_value("Kalkulation!J66")
    return max(0, h66 - j66)

def fkt_kalkulation_l66():
    n = get_excel_global('n')
    t = get_excel_global('t')
    VS = get_excel_global('VS')
    A66 = get_cell_value("Kalkulation!A66")
    H66 = get_cell_value("Kalkulation!H66")
    G66 = get_cell_value("Kalkulation!G66")
    try:
        if A66 > n:
            return 0
        elif A66 < t:
            return H66 / G66
        else:
            return VS
    except:
        return 0