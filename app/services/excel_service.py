import os
import pandas as pd


def load_excel(filepath: str) -> pd.DataFrame:
    if not filepath or not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")
    return pd.read_excel(filepath)


def detect_numeric_columns(df: pd.DataFrame) -> list[str]:
    numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
    # Exclude derived/aggregate columns from reuse in calculations
    excluded = {"total", "average", "sum"}
    numeric_cols = [col for col in numeric_cols if col.lower() not in excluded]
    return numeric_cols