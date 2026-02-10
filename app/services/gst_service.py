import pandas as pd
import re
from typing import Dict, Optional, Tuple, Union
from pathlib import Path
import io


def load_gst_sheet(filepath: Union[str, object]) -> pd.DataFrame:
    """
    Load GST sheet from Excel file or UploadedFile object

    Args:
        filepath: Path to the GST Excel file or UploadedFile object

    Returns:
        DataFrame with GST mapping data
    """
    if not filepath:
        return pd.DataFrame()

    try:
        # Handle UploadedFile object (from Streamlit)
        if hasattr(filepath, 'read'):
            # This is an UploadedFile object
            gst_df = pd.read_excel(filepath, sheet_name=0)
        else:
            # This is a file path string
            if not Path(filepath).exists():
                return pd.DataFrame()
            gst_df = pd.read_excel(filepath, sheet_name=0)

        # Auto-detect columns
        gst_mapping = auto_detect_gst_columns(gst_df)

        if not gst_mapping.get("item_desc") or not gst_mapping.get("tax_slab"):
            raise ValueError("Could not detect required columns in GST sheet")

        # Clean and normalize the data
        result_df = pd.DataFrame({
            "ItemDesc": gst_df[gst_mapping["item_desc"]].fillna("").astype(str).str.strip(),
            # Keep raw values so we can correctly parse formats like "18%" or 0.18 later
            "TaxSlab": gst_df[gst_mapping["tax_slab"]].fillna("").astype(str).str.strip(),
        })

        # Remove rows with empty item descriptions
        result_df = result_df[result_df["ItemDesc"] != ""].copy()

        return result_df

    except Exception as e:
        print(f"Error loading GST sheet: {e}")
        return pd.DataFrame()


def auto_detect_gst_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Auto-detect GST sheet columns

    Args:
        df: DataFrame to analyze

    Returns:
        Dictionary with detected column mappings
    """
    def _clean(s: str) -> str:
        if s is None:
            return ""
        s = str(s).strip().lower()
        s = re.sub(r'[_\-]+', ' ', s)
        s = re.sub(r'\s+', ' ', s)
        return s

    def _pick_column(df: pd.DataFrame, candidates):
        cleaned_to_original = {_clean(c): c for c in df.columns}

        for cand in candidates:
            cand_clean = _clean(cand)
            if cand_clean in cleaned_to_original:
                return cleaned_to_original[cand_clean]
            for k, orig in cleaned_to_original.items():
                if cand_clean and cand_clean in k:
                    return orig
        return ""

    result = {
        "item_desc": _pick_column(df, [
            "item desc", "item description", "description", "material",
            "material name", "item name", "item", "itemdesc", "item_desc"
        ]),
        "tax_slab": _pick_column(df, [
            "new tax slab (%)", "new tax slab", "tax slab (%)", "tax slab",
            "gst", "gst %", "tax %", "tax rate", "new tax slab %", "tax slab %"
        ])
    }

    # If no columns found, try to use the first two columns
    if not result["item_desc"] and not result["tax_slab"]:
        if len(df.columns) >= 2:
            result["item_desc"] = df.columns[1]  # Second column (Item Desc)
            result["tax_slab"] = df.columns[4] if len(df.columns) > 4 else df.columns[-1]  # Last column (Tax Slab)

    return result


def _normalize_item_desc(value: str) -> str:
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\u00a0", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = s.replace("/", " ")
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().upper()


def create_gst_lookup(gst_df: pd.DataFrame) -> Dict[str, float]:
    """
    Create GST lookup dictionary from GST DataFrame

    Args:
        gst_df: DataFrame with ItemDesc and TaxSlab columns

    Returns:
        Dictionary mapping item descriptions to tax slabs
    """
    if gst_df.empty:
        return {}

    # Create lookup dictionary
    gst_lookup = {}
    for _, row in gst_df.iterrows():
        item_desc = _normalize_item_desc(row["ItemDesc"])
        tax_slab_raw = str(row["TaxSlab"]).strip()

        # Handle percentage values like "18%" or "18"
        try:
            if tax_slab_raw.endswith('%'):
                tax_slab = float(tax_slab_raw[:-1])  # Remove % and convert to float
            else:
                # If it's already a decimal like 0.18, convert to percentage
                tax_slab = float(tax_slab_raw)
                if tax_slab < 1:  # If less than 1, it's a decimal percentage
                    tax_slab = tax_slab * 100  # Convert 0.18 to 18
        except ValueError:
            tax_slab = 0.0

        if item_desc:
            gst_lookup[item_desc] = tax_slab

    return gst_lookup


def match_gst_slab(item_desc: str, gst_lookup: Dict[str, float]) -> float:
    """
    Match item description to GST slab using fast matching

    Args:
        item_desc: Item description to match
        gst_lookup: GST lookup dictionary

    Returns:
        GST slab percentage (0.0 if no match found)
    """
    if not item_desc or not gst_lookup:
        return 0.0

    # Clean and normalize the item description
    item_desc_clean = _normalize_item_desc(item_desc)

    # Try exact match first (fastest)
    if item_desc_clean in gst_lookup:
        return gst_lookup[item_desc_clean]

    # Try simple substring matching (faster than word-based)
    for gst_item, tax_slab in gst_lookup.items():
        if gst_item in item_desc_clean or item_desc_clean in gst_item:
            return tax_slab

    # Only try word-based matching if GST lookup is small (< 50 items)
    if len(gst_lookup) < 50:
        item_words = set(item_desc_clean.split())
        best_match_score = 0
        best_match_slab = 0.0

        for gst_item, tax_slab in gst_lookup.items():
            gst_words = set(gst_item.split())
            # Calculate match score
            common_words = item_words.intersection(gst_words)
            match_score = len(common_words) / max(len(item_words), 1)

            if match_score > best_match_score and match_score > 0.5:  # Higher threshold for speed
                best_match_score = match_score
                best_match_slab = tax_slab

        return best_match_slab

    return 0.0


def calculate_gst_amount(base_amount: float, gst_percentage: float) -> float:
    """
    Calculate GST amount from base amount and percentage

    Args:
        base_amount: Base amount before GST
        gst_percentage: GST percentage (e.g., 18.0 for 18%)

    Returns:
        GST amount
    """
    if base_amount <= 0 or gst_percentage <= 0:
        return 0.0

    # GST percentage is already in percentage form (18.0 = 18%)
    # So we divide by 100 to get the decimal form for calculation
    return (base_amount * gst_percentage) / 100.0


def add_gst_columns(df: pd.DataFrame, gst_lookup: Dict[str, float]) -> pd.DataFrame:
    """
    Add GST-related columns to DataFrame

    Args:
        df: DataFrame to modify
        gst_lookup: GST lookup dictionary

    Returns:
        DataFrame with added GST columns
    """
    result_df = df.copy()

    # Find the item description column in GIN data (could be "ItemDesc", "Item Desc", etc.)
    item_desc_col = None
    possible_names = ["ITEM DESC", "ItemDesc", "Item Desc", "item desc", "item_desc", "Item Description", "item description", "Material", "material"]

    for col_name in possible_names:
        if col_name in result_df.columns:
            item_desc_col = col_name
            break

    if item_desc_col is None:
        # If no item description column found, add empty GST columns
        result_df["GST Slab"] = 0.0
        result_df["GST Amount"] = 0.0
        result_df["Total ISSUE Amount with GST"] = result_df.get("Amount", 0)
        return result_df

    # Add GST Slab column with % symbol
    gst_slab_values = result_df[item_desc_col].apply(
        lambda x: match_gst_slab(x, gst_lookup)
    )

    # Convert to percentage format (e.g., 18 -> 18%)
    result_df["GST Slab"] = gst_slab_values.apply(
        lambda x: f"{x}%" if x > 0 else "0%"
    )

    # If all GST Slab values are 0, try a simple fallback
    if gst_slab_values.sum() == 0 and gst_lookup:
        # Try to match the first item manually
        first_item = result_df[item_desc_col].iloc[0] if len(result_df) > 0 else ""
        for gst_item, tax_slab in gst_lookup.items():
            if gst_item in first_item or first_item in gst_item:
                result_df["GST Slab"] = f"{tax_slab}%"
                gst_slab_values = pd.Series([tax_slab] * len(result_df))
                break

    # Find amount column once (faster than checking every row)
    amount_col = None
    amount_cols = ["Amount", "ISSUE AMOUNT", "Issue Amount", "amount", "issue amount", "AMOUNT"]
    for col in amount_cols:
        if col in result_df.columns:
            amount_col = col
            break

    if amount_col:
        # Fast vectorized calculation
        result_df["GST Amount"] = (result_df[amount_col] * gst_slab_values / 100).fillna(0)
    else:
        result_df["GST Amount"] = 0.0


    # Add Total Amount with GST column (fast vectorized calculation)
    if amount_col:
        result_df["Total ISSUE Amount with GST"] = result_df[amount_col] + result_df["GST Amount"]
    else:
        result_df["Total ISSUE Amount with GST"] = result_df["GST Amount"]

    return result_df


def process_gst_for_gin_mapped(gin_df: pd.DataFrame, gst_file: Optional[Union[str, object]] = None) -> pd.DataFrame:
    """
    Process GST for GIN Mapped DataFrame

    Args:
        gin_df: GIN Mapped DataFrame
        gst_file: Optional path to GST file or UploadedFile object

    Returns:
        DataFrame with GST columns added
    """
    if not gst_file:
        # If no GST file provided, add empty GST columns
        result_df = gin_df.copy()
        result_df["GST Slab"] = 0.0
        result_df["GST Amount"] = 0.0
        result_df["Total ISSUE Amount with GST"] = result_df.get("Amount", 0)
        return result_df

    # Load GST data
    gst_df = load_gst_sheet(gst_file)

    if gst_df.empty:
        # If GST file is empty or invalid, add empty GST columns
        result_df = gin_df.copy()
        result_df["GST Slab"] = 0.0
        result_df["GST Amount"] = 0.0
        result_df["Total ISSUE Amount with GST"] = result_df.get("Amount", 0)
        return result_df

    # Create GST lookup
    gst_lookup = create_gst_lookup(gst_df)

    # Simple test - if no GST lookup, add empty columns
    if not gst_lookup:
        result_df = gin_df.copy()
        result_df["GST Slab"] = 0.0
        result_df["GST Amount"] = 0.0
        result_df["Total ISSUE Amount with GST"] = result_df.get("Amount", 0)
        return result_df

    # Add GST columns
    result_df = add_gst_columns(gin_df, gst_lookup)

    return result_df
