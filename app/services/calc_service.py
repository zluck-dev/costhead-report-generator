import pandas as pd


def add_sum_column(df: pd.DataFrame, cols: list[str]) -> str:
    # Ensure we calculate from original numeric cols only; overwrite if exists
    if "Sum" in df.columns:
        df.drop(columns=["Sum"], inplace=True)
    df["Sum"] = df[cols].sum(axis=1)
    return "Sum"


def add_average_column(df: pd.DataFrame, cols: list[str]) -> str:
    if "Average" in df.columns:
        df.drop(columns=["Average"], inplace=True)
    df["Average"] = df[cols].mean(axis=1)
    return "Average"