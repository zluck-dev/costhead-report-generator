import pandas as pd
import re
from typing import List


def list_sheets(filepath: str) -> List[str]:
    xls = pd.ExcelFile(filepath)
    return xls.sheet_names


def load_sheet(filepath: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(filepath, sheet_name=sheet_name)


def _clean(s: str) -> str:
    """Clean column name for fuzzy matching - same as working script"""
    if s is None: return ""
    s = str(s).strip().lower()
    s = re.sub(r'[_\-]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s


def _pick_column(df: pd.DataFrame, candidates):
    """Pick column using fuzzy matching - same as working script"""
    print(f"Available columns: {list(df.columns)}")
    cleaned_to_original = { _clean(c): c for c in df.columns }
    print(f"Cleaned columns: {cleaned_to_original}")

    for cand in candidates:
        cand_clean = _clean(cand)
        print(f"Trying candidate: '{cand}' -> cleaned: '{cand_clean}'")
        if cand_clean in cleaned_to_original:
            result = cleaned_to_original[cand_clean]
            print(f"Exact match found: '{result}'")
            return result
        for k, orig in cleaned_to_original.items():
            if cand_clean and cand_clean in k:
                result = orig
                print(f"Partial match found: '{result}' (contains '{cand_clean}')")
                return result
    print("No match found")
    return ""


def auto_detect_columns(df: pd.DataFrame, sheet_type: str) -> dict:
    """Auto-detect columns using the same logic as working script"""
    print(f"\nAuto-detecting columns for {sheet_type} sheet...")
    if sheet_type == "master":
        result = {
            "code": _pick_column(df, ["activity code","activity_code","act code","actcode","code"]),
            "name": _pick_column(df, ["activity name","activity_name","act name","actname","name","task name","task"]),
            "wbs": _pick_column(df, ["parent wbs name","parent wbs","wbs name","wbs parent","wbs level 1","wbs","parent"])
        }
    else:  # gin
        result = {
            "code": _pick_column(df, ["activity code","activity_code","act code","actcode","code","act code in gin"])
        }
    print(f"Detection result: {result}")
    return result


def _normalize_code_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.replace("\u00A0", " ", regex=False)


def build_activity_lookup(df_activities: pd.DataFrame, code_col: str, wbs_col: str, name_col: str) -> pd.DataFrame:
    print(f"\nBuilding lookup with columns: code='{code_col}', wbs='{wbs_col}', name='{name_col}'")
    print(f"Activities DataFrame columns: {list(df_activities.columns)}")

    lookup = df_activities[[code_col, wbs_col, name_col]].copy()
    # Normalize codes to strings for safe join
    lookup[code_col] = _normalize_code_series(lookup[code_col])
    # Drop duplicates on code, keep first occurrence
    lookup = lookup.dropna(subset=[code_col]).drop_duplicates(subset=[code_col], keep="first")
    lookup = lookup.rename(columns={wbs_col: "ParentWBS", name_col: "ActivityName", code_col: "ActivityCode"})
    print(f"Lookup created with {len(lookup)} rows")
    print(f"Lookup columns: {list(lookup.columns)}")
    return lookup


def merge_gin_with_lookup(
    df_gin: pd.DataFrame,
    gin_code_col: str,
    activities_lookup: pd.DataFrame,
) -> pd.DataFrame:
    print(f"\nMerging GIN with lookup...")
    print(f"GIN DataFrame columns: {list(df_gin.columns)}")
    print(f"GIN code column: '{gin_code_col}'")
    print(f"Activities lookup columns: {list(activities_lookup.columns)}")

    gin = df_gin.copy()
    # Normalize code in GIN
    gin[gin_code_col] = _normalize_code_series(gin[gin_code_col])

    # Align code column name for merge
    left = gin.rename(columns={gin_code_col: "ActivityCode"})
    print(f"After rename, left columns: {list(left.columns)}")

    merged = left.merge(activities_lookup, on="ActivityCode", how="left")
    print(f"After merge, columns: {list(merged.columns)}")

    # Order so that matched rows appear first, unmatched at the end (like the working script)
    matched_mask = (~merged["ActivityName"].isna()) & (~merged["ParentWBS"].isna())
    merged["_matched_sort"] = matched_mask.astype(int)  # 1 for matched, 0 for unmatched
    merged = merged.sort_values(by=["_matched_sort"], ascending=False).drop(columns=["_matched_sort"])

    # Replace NaN with blanks (as requested) - this is the key fix from the working script
    merged["ActivityName"] = merged["ActivityName"].fillna("")
    merged["ParentWBS"] = merged["ParentWBS"].fillna("")

    # Reorder columns: original GIN columns first, then ActivityCode, ParentWBS, ActivityName, then GST columns at the end
    cols = list(merged.columns)
    print(f"Before reordering, columns: {cols}")

    # Remove the columns we want to move to the end
    gst_columns = []
    if "ActivityCode" in cols:
        cols.remove("ActivityCode")
    if "ParentWBS" in cols:
        cols.remove("ParentWBS")
    if "ActivityName" in cols:
        cols.remove("ActivityName")

    # Check for GST columns and remove them from the main list
    gst_col_names = ["GST Slab", "GST Amount", "Total ISSUE Amount with GST"]
    for gst_col in gst_col_names:
        if gst_col in cols:
            cols.remove(gst_col)
            gst_columns.append(gst_col)

    # Add the mapped columns at the end, then GST columns
    cols = cols + ["ActivityCode", "ParentWBS", "ActivityName"] + gst_columns
    print(f"After reordering, columns: {cols}")
    merged = merged[cols]
    return merged
