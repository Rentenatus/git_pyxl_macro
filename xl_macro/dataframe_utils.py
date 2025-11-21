# -*- coding: utf-8 -*-
import pandas as pd


def chunks_to_dataframe(chunks) -> pd.DataFrame:
    """
    Wandelt die extrahierten Code-Chunks in ein Pandas DataFrame um.

    # Bsp.:
    for idx, row in df.iterrows():
        # Beispiel: Kombiniere meaning und line_start zu einem Kommentar
        python_code = f"# {row['meaning']} ab Zeile {row['line_start']}\n{row['code']}"
        df.at[idx, 'python_code'] = python_code

    :param chunks:
    :return: DataFrame
    """
    df = pd.DataFrame(chunks, columns=[
        "meaning", "params", "code", "line_start", "line_end", "local_used"
    ])
    df["python_code"] = ""  # Leere Spalte ergänzen
    return df


def save_dataframe_as(df: pd.DataFrame, filename: str):
    """
    Saves the DataFrame, format DataFrame, and structure to disk.
    """
    try:
        import json
    except ImportError:
        raise ImportError("The package 'json' is not installed. Install it with 'pip install json'.")

    # Values als CSV
    df.to_csv(f"{filename}_df.csv", index=False)

    # Structure als JSON
    structure = {col: str(dtype) for col, dtype in df.dtypes.items()}
    with open(f"{filename}_st.json", "w", encoding="utf-8") as f:
        json.dump(structure, f, indent=2)


def load_dataframe(filename: str) -> pd.DataFrame:
    """
    Loads the DataFrame, and structure from disk.
    Returns a new TableEnginePandas instance.
    """
    try:
        import json
    except ImportError:
        raise ImportError("The package 'json' is not installed. Install it with 'pip install json'.")

    # Structure
    with open(f"{filename}_st.json", "r", encoding="utf-8") as f:
        structure = json.load(f)

    # Values
    return pd.read_csv(f"{filename}_df.csv", dtype=structure)

