import pandas as pd
import re
import warnings
from typing import List, Dict, Tuple
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from .gst_service import process_gst_for_gin_mapped

warnings.filterwarnings("ignore")

# ---------- Synonyms ----------
GIN_SYNS = {
    "ActivityName": ["activity name","activity_name","task name","task","act name"],
    "ParentWBS":    ["parent wbs","wbs name","parent wbs name","wbs parent","wbs level 1","wbs"],
    "SubProject":   ["sub project","subproject","sub_project","location","subproject name","sub proj","sub_proj"],
    "Project":      ["project","project name","project_name"],
    "Amount":       ["gin issue amt","issue amount","amount","total issued amount","issued amount","value","net amount","total amount"],
    "IssueQty":     ["issued quantity","issue qty","qty","quantity","issue quantity","issued qty"],
    "ItemGroup":    ["item group","itemgroup","group","material group","category"],
    "ItemDesc":     ["item desc","item description","description","material","material name","item name","item"],
    "Remarks":      ["remarks","note","notes","narration","comment"],
}

MAP_COMPOSITE_SYNS = {
    "CostHead":   ["cost heads","cost head","costhead","head"],
    "Value":      ["parent wbs / activity name / sub project","parent wbs/activity name/sub project","value","criteria","keyword"],
    "FromWhere":  ["from where","match in","field","where"],
}

# ---------- Special keyword classification (report_simple_v7 logic) ----------
EXCLUDE_ANY_DESC = [
    "CPVC PIPE", "UPVC PIPE",
    "ASHIRVAD CPVC", "ASHIRVAD UPVC", "BIRLA UPVC",
    "NEGEV SAND MATT", "GVT TILE", "THREAD SAND TEXTURE", "8'X4' TILES"
]
EXCLUDE_ANY_GROUP = ["CPVC", "UPVC"]

# Excluded keywords for ParentWBS fallback (unmatched items)
# If any of these appear in ItemGroup/ItemDesc/Remarks, do NOT apply
# ParentWBS-based fallbacks; instead list them under a special section
# in the summary ("Unmatched Item Group & Items").
EXCLUDED_UNMATCHED_KEYWORDS_DESC = [
    # Add business-specific keywords here, examples:
    # "DOOR FRAME", "WINDOW FRAME", "GLAZING"
]
EXCLUDED_UNMATCHED_KEYWORDS_GROUP = [
    # Add group-level keywords here if needed
    "BORING","WATERPROOING CHEMICALS","SUPREME AGRI","GI A CLASS","CEMENT SHEET",
    "GI - A CLASS","GI B CLASS","GI - B CLASS","WATERPROOING CHEMICALS","TILE ADHESIVE","M.S. PLATE","LOAD BEARING PAD",
]

STEEL_KEYWORDS_DESC = [
    "TMT","REBAR","REINFORCEMENT","MS ROD","TOR STEEL","BINDING WIRE","STEEL BAR",
    "TMT BAR","DOWEL","CHAIR","REO","DEFORMED BAR","RIBBED BAR"
]
STEEL_KEYWORDS_GROUP = ["STEEL"]
STEEL_EXCLUDE_DESC = ["PPC","OPC","CEMENT","C CHANNEL","CHANNEL","MS C CHANNEL","M.S. C - CHANNEL"]
STEEL_EXCLUDE_GROUP = ["CEMENT","PPC","OPC","CHANNEL","REBAR CHEMICAL"]

# Optional excludes for Concrete and Masonry classification (similar to Steel)
CONCRETE_EXCLUDE_DESC = [
    # e.g., "MASONRY", "BLOCK", "BRICK"
]
CONCRETE_EXCLUDE_GROUP = [
    # e.g., "MASONRY"
]

# Exclude lists for Masonry classification
MASONRY_EXCLUDE_DESC = [
    # e.g., "CONCRETE", "RMC"
]
MASONRY_EXCLUDE_GROUP = [
    # e.g., "CONCRETE", "RMC"
    "SUPREME AGRI"
]

CONCRETE_KEYWORDS_DESC = [
    "RMC","READY MIX","READY-MIX","READYMIX","TRANSIT MIX","PUMPED CONCRETE",
    "M20","M25","M30","M35","M40","DESIGN MIX","SITE MIX CONCRETE",
    "FLY ASH CONCRETE","GGBS CONCRETE","WITH FLY ASH","WITHOUT FLY ASH","GGBS","GGBFS",
    "OPC CEMENT","OPC","ORDINARY PORTLAND CEMENT",
    "FLY ASH",
    "SLAG 80-100MM","SLAG 40-80MM","SLAG",
    "6MM MAMRI","MAMRI 6MM",
    "GSB",
    "20MM KAPCHI","KAPCHI 20MM",
    "RUBBLE"
]
CONCRETE_KEYWORDS_GROUP = ["CONCRETE","CEMENT CONCRETE","RMC","READY MIX","OPC","WATERPROOFING HARDWARE"]

MASONRY_KEYWORDS_DESC = [
    "RIVER SAND","POICHA","M SAND","MSAND","SAND",
    "PPC CEMENT","CEMENT PPC","PPC (BAG)","PPC BAG",
    "AAC BLOCK","AUTOCLAVED AERATED",
    "SBR","SIKA LATEX",
    "FIBER CHICKEN MESH","CHICKEN MESH","FIBER MESH",
    "BLOCK JOINT MORTAR","BLOCK JOINTING MORTAR","TBM BLOCK JOINTING MORTAR","BLOCK FIX","JOINING MORTAR",
    "FLYASH BRICK","FLY ASH BRICK","RED BRICK","REGULAR RED BRICKS","BRICK",
    "SIKA GROUT 214 1N","SIKA GROUT 214","GROUT 214","GROUT",
    "NON ISI PVC PIPE","PVC PIPE"
]
MASONRY_KEYWORDS_GROUP = ["AAC","BRICK","BLOCK","MASONRY","PLASTER","PPC","SAND","PVC PIPE","NON ISI PVC",
                          "REBAR CHEMICAL","NON ISI PLUMBING","REGULAR HARDWARE"]

# ---------- Helpers ----------
def _clean_name(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().lower()
    s = re.sub(r'[_\-]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s

def _norm_text(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = s.replace("\u00a0"," ")
    s = s.replace("–","-").replace("—","-")
    s = s.replace("/", " ")
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().upper()

def _pick_col(df: pd.DataFrame, options):
    clean = {_clean_name(c): c for c in df.columns}
    for cand in options:
        k = _clean_name(cand)
        if k in clean: return clean[k]
        for kk, orig in clean.items():
            if k in kk: return orig
    return ""

def _pick_cols(df: pd.DataFrame, spec: dict):
    return {k: _pick_col(df, [k]+v) for k,v in spec.items()}

def _detect_header(df_raw: pd.DataFrame, key_words=("cost","head","activity","wbs","sub","project","amount","qty","item"), scan=20):
    df = df_raw.copy()
    if len(df) == 0: return df
    upto = min(scan, len(df))
    for r in range(upto):
        row = [str(x).strip().lower() for x in df.iloc[r].tolist()]
        if any(k in " ".join(row) for k in key_words):
            df.columns = df.iloc[r]
            return df[r+1:].reset_index(drop=True)
    df.columns = df.iloc[0]
    return df[1:].reset_index(drop=True)

def _norm_series_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).map(_norm_text)

def _norm_series_amt(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.replace(",", "", regex=False).str.replace("\u00a0", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _split_keywords(v: str):
    # Improve keyword splitting: do NOT split by commas globally because values
    # often contain descriptive parentheses with commas. We split by ; or | only.
    raw = (v or "")
    base_parts = re.split(r"[;\|]", raw)
    keywords: List[str] = []
    for part in base_parts:
        p = part.strip()
        if not p:
            continue
        # full text as keyword
        kw_full = _norm_text(p)
        if kw_full and kw_full not in keywords:
            keywords.append(kw_full)
        # also add text before first parenthesis as a broader keyword
        m = re.match(r"^(.*?)\s*\(", p)
        if m:
            base = _norm_text(m.group(1))
            if base and base not in keywords:
                keywords.append(base)
    return keywords

def _field_for(where_text: str) -> str:
    w = (where_text or "").lower()
    if "activity" in w or "task" in w: return "ActivityName"
    if "wbs" in w:                      return "ParentWBS"
    if "sub" in w:                      return "SubProject"
    if "project" in w:                  return "Project"
    if "itemgroup" in w or "item group" in w: return "ItemGroup"
    if "itemdesc" in w or "item desc" in w: return "ItemDesc"
    if "remarks" in w:                   return "Remarks"
    return "ActivityName"

# ---------- SubProject Filter ----------
def _norm_sub(s: str) -> str:
    if s is None: return ""
    s = str(s).upper().replace('-', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def _is_allowed_subproject(sp: str) -> bool:
    spn = _norm_sub(sp)
    if spn == "PODIUM" or spn == "NON TOWER AREA":
        return True
    if spn.startswith("TOWER ") and len(spn) == len("TOWER X"):
        letter = spn.split(" ")[-1]
        return len(letter) == 1 and 'A' <= letter <= 'Z'
    return False

# ---------- Special Heads Classifier ----------
def _contains_exact_keyword(text: str, keywords: List[str]) -> bool:
    """Check if text contains any of the keywords as complete words (not substrings)"""
    if not text or not keywords:
        return False
    text_upper = text.upper()
    for keyword in keywords:
        keyword_upper = keyword.upper()
        # For keywords with special characters like parentheses, use exact matching
        if '(' in keyword_upper or ')' in keyword_upper:
            if keyword_upper in text_upper:
                return True
        else:
            # Use word boundaries for simple keywords
            import re
            pattern = r'\b' + re.escape(keyword_upper) + r'\b'
            if re.search(pattern, text_upper):
                return True
    return False

def _is_globally_excluded(item_group: str, item_desc: str, remarks: str) -> bool:
    G, D, R = (str(item_group or "").upper(), str(item_desc or "").upper(), str(remarks or "").upper())
    if _contains_exact_keyword(D, EXCLUDE_ANY_DESC) or _contains_exact_keyword(R, EXCLUDE_ANY_DESC):
        return True
    if _contains_exact_keyword(G, EXCLUDE_ANY_GROUP):
        return True
    return False

def _is_excluded_for_unmatched(item_group: str, item_desc: str, remarks: str) -> bool:
    """Check if a row should be excluded from ParentWBS fallback due to keywords."""
    G = str(item_group or "").upper()
    D = str(item_desc or "").upper()
    R = str(remarks or "").upper()
    if _contains_exact_keyword(D, EXCLUDED_UNMATCHED_KEYWORDS_DESC):
        return True
    if _contains_exact_keyword(R, EXCLUDED_UNMATCHED_KEYWORDS_DESC):
        return True
    if _contains_exact_keyword(G, EXCLUDED_UNMATCHED_KEYWORDS_GROUP):
        return True
    return False

def _classify_special_head(row: Dict) -> str:
    g = str(row.get("ItemGroup","") or "").upper()
    d = str(row.get("ItemDesc","") or "").upper()
    r = str(row.get("Remarks","") or "").upper()
    hay_desc = f"{d} {r}".strip()
    hay_group = g
    # Only extract Steel/Concrete/Masonry for allowed SubProjects (PODIUM or TOWER A-Z)
    sp_allowed = _is_allowed_subproject(row.get("SubProject", ""))

    if _is_globally_excluded(g, d, r):
        return "Other"

    steel_hit = _contains_exact_keyword(hay_desc, STEEL_KEYWORDS_DESC) or _contains_exact_keyword(hay_group, STEEL_KEYWORDS_GROUP)
    if steel_hit:
        if sp_allowed and not (_contains_exact_keyword(hay_desc, STEEL_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, STEEL_EXCLUDE_GROUP)):
            return "Steel"

    if sp_allowed and (_contains_exact_keyword(hay_desc, MASONRY_KEYWORDS_DESC) or _contains_exact_keyword(hay_group, MASONRY_KEYWORDS_GROUP)):
        # Respect Masonry excludes
        if not (_contains_exact_keyword(hay_desc, MASONRY_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, MASONRY_EXCLUDE_GROUP)):
            return "Masonry and plaster material only"

    if sp_allowed:
        # Respect Concrete excludes
        for kw in CONCRETE_KEYWORDS_DESC:
            if _contains_exact_keyword(hay_desc, [kw]) and not (_contains_exact_keyword(hay_desc, CONCRETE_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, CONCRETE_EXCLUDE_GROUP)):
                return "Concrete"
        for kw in CONCRETE_KEYWORDS_GROUP:
            if _contains_exact_keyword(hay_group, [kw]) and not (_contains_exact_keyword(hay_desc, CONCRETE_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, CONCRETE_EXCLUDE_GROUP)):
                return "Concrete"

    return "Other"

# ---------- Canonicalize CostHead labels ----------
def _canonicalize_costhead(value: str) -> str:
    s = (value or "").strip()
    key = re.sub(r"[^a-z]+", " ", s.lower()).strip()
    if key == "steel":
        return "Steel"
    if key == "concrete":
        return "Concrete"
    if key in {"masonry and plaster material only", "masonry plaster material only", "masonry and plaster"}:
        return "Masonry and plaster material only"
    return s

# ---------- Logging Utils ----------
def _setup_logger(log_file: Path):
    pass

# ---------- Loaders ----------
def load_gin(df_raw: pd.DataFrame) -> pd.DataFrame:
    gin = _detect_header(df_raw, key_words=("activity","wbs","sub","project","amount","issue","qty","item"))
    cols = _pick_cols(gin, GIN_SYNS)
    need = ["ActivityName","ParentWBS","SubProject","Amount"]
    miss = [c for c in need if not cols[c]]
    if not cols.get("Project"):
        gin["__ProjectBlank__"] = ""
        cols["Project"] = "__ProjectBlank__"
    if miss:
        raise ValueError(f"GIN_Mapped missing columns: {miss}\nFound: {list(gin.columns)}")

    out = pd.DataFrame({
        "ActivityName": _norm_series_text(gin[cols["ActivityName"]]),
        "ParentWBS":    _norm_series_text(gin[cols["ParentWBS"]]),
        "SubProject":   _norm_series_text(gin[cols["SubProject"]]),
        "Project":      _norm_series_text(gin[cols["Project"]]),
        "Amount":       _norm_series_amt(gin[cols["Amount"]]),
    })
    out["ItemGroup"] = _norm_series_text(gin[cols["ItemGroup"]]) if cols.get("ItemGroup") else ""
    out["ItemDesc"]  = _norm_series_text(gin[cols["ItemDesc"]])  if cols.get("ItemDesc")  else ""
    out["IssueQty"]  = _norm_series_amt(gin[cols["IssueQty"]])   if cols.get("IssueQty")  else 0.0
    out["Remarks"]   = _norm_series_text(gin[cols["Remarks"]])   if cols.get("Remarks")   else ""

    return out

def load_mapping(df_raw: pd.DataFrame) -> pd.DataFrame:
    mp = _detect_header(df_raw, key_words=("cost","head","activity","wbs","sub","from","project"))
    comp = _pick_cols(mp, MAP_COMPOSITE_SYNS)
    if not comp.get("CostHead") or not comp.get("Value") or not comp.get("FromWhere"):
        raise ValueError("CostHead sheet not recognized. Expected: COST HEADS | parent WBS / Activity Name / Sub Project | From where")
    df = pd.DataFrame({
        "CostHead":  mp[comp["CostHead"]].fillna("").astype(str).map(_norm_text),
        "ValueRaw":  mp[comp["Value"]].fillna("").astype(str),
        "FromWhere": mp[comp["FromWhere"]].fillna("").astype(str),
        "Value":     mp[comp["Value"]].fillna("").astype(str).map(_norm_text),
    })
    return df

# ---------- Mapping ----------
def assign_cost_head(gin: pd.DataFrame, mapping: pd.DataFrame, match_mode="contains"):
    tagged = gin.copy()
    tagged["CostHead"] = "Other"
    assigned_mask = pd.Series(False, index=tagged.index)

    # FIRST: Apply special heads (Steel, Concrete, Masonry) based on keywords
    specials = tagged.apply(_classify_special_head, axis=1)
    special_mask = specials.isin(["Steel","Concrete","Masonry and plaster material only"])
    tagged.loc[special_mask, "CostHead"] = specials[special_mask]
    assigned_mask = assigned_mask | special_mask

    # SECOND: Apply mapping rules only to remaining unassigned rows
    expanded = []
    special_heads = {"Steel", "Concrete", "Masonry and plaster material only"}
    for _, r in mapping.iterrows():
        head_raw = r["CostHead"]
        head_can = _canonicalize_costhead(head_raw)
        # Skip any mapping rule that tries to set a special head; those are keyword-only
        if head_can in special_heads:
            continue
        # Support multi-field FromWhere like "ParentWBS, ActivityName"
        fields_raw = str(r["FromWhere"] or "")
        fields = [f.strip() for f in re.split(r"[;,]", fields_raw) if f.strip()]
        if not fields:
            fields = [_field_for(r["FromWhere"])]
        for kw in _split_keywords(r["Value"]):
            if head_raw and kw:
                for f in fields:
                    field = _field_for(f)
                    expanded.append({"CostHead": head_raw, "Field": field, "Keyword": kw})
    map_expanded = pd.DataFrame(expanded)

    # Apply mapping rules only to unassigned rows
    assigned_by_mapping = 0
    if not map_expanded.empty:
        for _, mr in map_expanded.iterrows():
            field = mr["Field"]; kw = mr["Keyword"]; hay = tagged[field]
            if match_mode == "exact":
                mask = (hay == kw)
            else:
                try:
                    mask = hay.str.contains(re.escape(kw), na=False)
                except Exception:
                    mask = hay.fillna("").astype(str).str.contains(kw, na=False)
            to_assign = mask & (~assigned_mask)
            if to_assign.any():
                tagged.loc[to_assign, "CostHead"] = mr["CostHead"]
                assigned_mask = assigned_mask | to_assign
                assigned_now = int(to_assign.sum())
                assigned_by_mapping += assigned_now

    # Return after applying specials first, then mapping
    return tagged, map_expanded

# ---------- Reallocate unmatched by SubProject ----------
def reallocate_unmatched_by_subproject(tagged: pd.DataFrame) -> pd.DataFrame:
    t = tagged.copy()
    # Disabled per user request: do not introduce extra CostHeads based on SubProject
    return t

# ---------- Fallbacks based on ParentWBS for unmatched ----------
def apply_parentwbs_fallbacks(tagged: pd.DataFrame) -> pd.DataFrame:
    t = tagged.copy()
    # Track rows excluded from fallback by excluded keywords
    t["__ExcludedByKeywords__"] = False
    is_other = (t["CostHead"] == "Other")
    # Identify excluded unmatched rows first: match excluded keywords AND has Activity
    has_activity = t["ActivityName"].fillna("") != ""
    excl_mask = is_other & has_activity & t.apply(
        lambda r: _is_excluded_for_unmatched(r.get("ItemGroup"), r.get("ItemDesc"), r.get("Remarks")), axis=1
    )
    if excl_mask.any():
        t.loc[excl_mask, "__ExcludedByKeywords__"] = True

    # Only apply fallbacks to those not excluded
    candidates = is_other & (~t["__ExcludedByKeywords__"])
    wbs_norm = t["ParentWBS"].fillna("").astype(str).map(_norm_text)
    mask_masonry = candidates & (
        wbs_norm.str.contains(r"\bMASONARY WORK\b", regex=True) |
        wbs_norm.str.contains(r"\bINTERNAL PLASTER WORK\b", regex=True)
    )
    mask_concrete = candidates & wbs_norm.str.contains(r"\bRCC WORK\b", regex=True)
    t.loc[mask_masonry, "CostHead"] = "Masonry and plaster material only"
    t.loc[mask_concrete, "CostHead"] = "Concrete"
    return t

# ---------- Main CostHead Report Function ----------
def generate_costhead_report(gin_filepath: str, costhead_filepath: str, output_dir: str, match_mode: str = "contains", gst_filepath = None) -> Dict[str, str]:
    """
    Generate CostHead reports based on the build_costhead_report_v9.py script

    Args:
        gin_filepath: Path to the GIN_Mapped Excel file
        costhead_filepath: Path to the CostHead Excel file
        output_dir: Directory to save the output reports
        match_mode: "contains" or "exact" for matching mode
        gst_filepath: Optional path to GST Excel file or UploadedFile object

    Returns:
        Dict with paths to generated files
    """

    # Prepare logging/output early so we can trace the whole flow
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # Load the Excel files
    gin_xl = pd.ExcelFile(gin_filepath, engine="openpyxl")
    costhead_xl = pd.ExcelFile(costhead_filepath, engine="openpyxl")

    # Find the GIN_Mapped sheet (should be the main sheet)
    gin_sheet = None
    for sheet in gin_xl.sheet_names:
        if "gin" in sheet.lower() and "mapped" in sheet.lower():
            gin_sheet = sheet
            break
    if not gin_sheet:
        gin_sheet = gin_xl.sheet_names[0]  # Use first sheet if not found

    # Load raw data
    gin_raw = gin_xl.parse(gin_sheet, header=None)
    costhead_raw = costhead_xl.parse(costhead_xl.sheet_names[0], header=None)  # Use first sheet

    # Process the data
    gin = load_gin(gin_raw)

    # Process GST if GST file is provided
    if gst_filepath:
        gin = process_gst_for_gin_mapped(gin, gst_filepath)

    # Removed global SubProject filtering; filtering is enforced only for Steel/Concrete/Masonry within classifier

    mapping = load_mapping(costhead_raw)
    mapping = mapping[mapping["CostHead"] != ""].copy()

    # 1) Apply mapping (with special heads pre-assigned)
    tagged, map_expanded = assign_cost_head(gin, mapping, match_mode=match_mode)

    # Snapshot BEFORE fallbacks for unmatched views (so details include rows pre-fallback)
    tagged_before_fallbacks = tagged.copy()

    # 1b) Apply ParentWBS-based fallbacks so these are not counted as unmatched in final allocation
    tagged = apply_parentwbs_fallbacks(tagged)

    # 2) Reallocate any remaining "Other" by SubProject rules
    tagged_final = reallocate_unmatched_by_subproject(tagged)

    # Canonicalize CostHead labels to avoid duplicates like STEEL vs Steel
    tagged_final["CostHead"] = tagged_final["CostHead"].apply(_canonicalize_costhead)


    # 3) Build reports from the final classification
    breakdown = (
        tagged_final.groupby(["CostHead","ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["CostHead","ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc"])
    )

    matched_final = tagged_final.copy()  # everything, since we've allocated
    amenities_mask_summary = (
        (matched_final["CostHead"] == "Other") &
        ((matched_final["ActivityName"].fillna("") == "") | (matched_final["ParentWBS"].fillna("") == ""))
    )
    # Check if GST columns exist in the data
    has_gst_data = "GST Amount" in matched_final.columns and "GST Slab" in matched_final.columns

    if has_gst_data:
        # Include GST data in summary aggregation
        summary = (
            matched_final[~amenities_mask_summary]
            .groupby("CostHead", as_index=False)
            .agg({
                "Amount": "sum",
                "GST Amount": "sum"
            })
            .rename(columns={"Amount": "Material TotalAmount"})
            .sort_values("CostHead")
        )
        # Calculate total with GST
        summary["Total Material Amount with GST"] = summary["Material TotalAmount"] + summary["GST Amount"]
    else:
        # Fallback to original behavior if no GST data
        summary = (
            matched_final[~amenities_mask_summary]
            .groupby("CostHead", as_index=False)["Amount"]
            .sum().rename(columns={"Amount":"Material TotalAmount"})
            .sort_values("CostHead")
        )
        summary["GST Amount"] = 0.0
        summary["Total Material Amount with GST"] = summary["Material TotalAmount"]

    # Build ITEM column: aggregated unique tokens "ItemGroup | ItemDesc" per CostHead
    def _token(row):
        ig = str(row.get("ItemGroup", "") or "").strip()
        idc = str(row.get("ItemDesc", "") or "").strip()
        if ig and idc:
            return f"{ig} | {idc}"
        return idc or ig or "(blank)"

    def _token_with_gst(row):
        """Create token with GST information if available"""
        base_token = _token(row)
        if has_gst_data and "GST Slab" in row and pd.notna(row.get("GST Slab")):
            gst_slab = str(row.get("GST Slab", "")).strip()
            if gst_slab and gst_slab != "0%":
                return f"{base_token} - {gst_slab}"
        return base_token

    def _gst_item_token(row):
        """Create GST item token with only item description and GST percentage"""
        if has_gst_data and "GST Slab" in row and pd.notna(row.get("GST Slab")):
            gst_slab = str(row.get("GST Slab", "")).strip()
            if gst_slab and gst_slab != "0%":
                # Get just the item description without group
                idc = str(row.get("ItemDesc", "") or "").strip()
                if idc:
                    return f"- {idc} - {gst_slab}"
        return ""

    item_text: Dict[str, str] = {}
    gst_items_text: Dict[str, str] = {}
    allowed_item_heads = {"Steel", "Concrete", "Masonry and plaster material only"}

    # Process all cost heads for GST items, but only special heads for ITEM Remark
    for head in summary["CostHead"].tolist():
        subset = matched_final[matched_final["CostHead"] == head]

        # Process ITEM Remark only for special heads (without GST info)
        if head in allowed_item_heads:
            uniq = sorted(set(_token(r) for _, r in subset.iterrows()))

            MAX_ITEMS = 200
            if len(uniq) > MAX_ITEMS:
                shown = uniq[:MAX_ITEMS]
                more = len(uniq) - MAX_ITEMS
                item_text[head] = "; ".join(shown) + f"; ... (+{more} more)"
            else:
                item_text[head] = "; ".join(uniq)

        # Process GST items for ALL cost heads
        if has_gst_data:
            gst_items = [item for item in [_gst_item_token(r) for _, r in subset.iterrows()] if item]
            gst_items_uniq = sorted(set(gst_items))
            if gst_items_uniq:
                gst_items_text[head] = "\n".join(gst_items_uniq)
            else:
                gst_items_text[head] = ""
        else:
            gst_items_text[head] = ""

    summary["ITEM Remark"] = summary["CostHead"].map(item_text).fillna("")
    summary["GST Items"] = summary["CostHead"].map(gst_items_text).fillna("")

    # Build Indoor Amenities breakdown into ITEM cell as bullet list of SubProject totals
    indoor_item_excel_row = None
    try:
        indoor_mask_data = matched_final["CostHead"].astype(str).str.strip().str.lower() == "indoor amenities"
        if indoor_mask_data.any():
            sp_totals = (
                matched_final[indoor_mask_data]
                .groupby("SubProject", as_index=False)["Amount"].sum()
                .sort_values("SubProject")
            )
            # Clean names and build bullet lines
            lines = []
            for _, r in sp_totals.iterrows():
                sp = str(r["SubProject"]).strip()
                amt = float(r["Amount"]) if pd.notnull(r["Amount"]) else 0.0
                if sp:
                    lines.append(f"- {sp} - {amt:,.2f} Rs.")
            text = "\n".join(lines)
            # Write into summary ITEM cell for Indoor Amenities
            mask_summary = summary["CostHead"].astype(str).str.strip().str.lower() == "indoor amenities"
            if mask_summary.any():
                summary.loc[mask_summary, "ITEM Remark"] = text
    except Exception:
        pass

    # Enforce fixed CostHead ordering and merge specified variants
    desired_order = [
        "Excavation",
        "D--wall",
        "Rcc + masonry + plaster labour",
        "Steel",
        "Concrete",
        "Masonry and plaster material only",
        "UNIT FINISH (COST REQUIRE TO FINISH ONE. UNIT (FLAT SHOP OFFICE SERVANT ROOM) . EXCLUDING MASONRY & PLASTER",
        "Window section",
        "RAILING, GRILL",
        "STAIR, TYPICAL PASSAGE, BASEMENT PASSAGE, BASEMENT FOYER",
        "TERRACE FINISHING (WATERPROOFING AND FINISHING)",
        "Basement finishing only (including floor finishing, color and lighting, art work & excluding fire, basement exhaust)",
        "COMMON PLUMBING (INCLUDING PUMPS)",
        "Common electric (geb, dg, tc to meter room and meter room to flat mcb)",
        "Fire (including basement exhaust)",
        "LANDSCAPE, OUTDOOR AMENITIES, GROUND FLOOR DRIVEWAY AND PARKING, COMPOUND WALL & GATE",
        "Indoor amenities",
        "Louvers",
        "Outer paint and texture",
        "Elevation treatment (cladding)",
        "Lift",
        "Consultants",
        "Overhead",
        "MISCELLANOUS",
    ]

    combined_16 = "LANDSCAPE, OUTDOOR AMENITIES, GROUND FLOOR DRIVEWAY AND PARKING, COMPOUND WALL & GATE"

    # Build normalization map for desired labels and include 16's variants
    def _normalize_head_label(label: str) -> str:
        return _norm_text(label or "")

    desired_norm_map = { _normalize_head_label(h): h for h in desired_order }
    variant_norm_map = {
        _normalize_head_label("LANDSCAPE, OUTDOOR AMENITIES, GROUND FLOOR DRIVEWAY AND PARKING"): combined_16,
        _normalize_head_label("COMPOUND WALL & GATE"): combined_16,
    }

    def _merge_and_canonicalize_costhead(label: str) -> str:
        norm = _normalize_head_label(label)
        if norm in variant_norm_map:
            return variant_norm_map[norm]
        if norm in desired_norm_map:
            return desired_norm_map[norm]
        return str(label or "")

    # Apply merging/canonicalization on summary then re-aggregate in case of merges
    if not summary.empty:
        summary["CostHead"] = summary["CostHead"].map(_merge_and_canonicalize_costhead)
        if "Material TotalAmount" in summary.columns:
            agg_columns = {
                "Material TotalAmount": "sum",
                "ITEM Remark": lambda s: "; ".join([x for x in s if str(x).strip()])
            }
            if "GST Amount" in summary.columns:
                agg_columns["GST Amount"] = "sum"
            if "Total Material Amount with GST" in summary.columns:
                agg_columns["Total Material Amount with GST"] = "sum"
            if "GST Items" in summary.columns:
                agg_columns["GST Items"] = lambda s: "\n".join([x for x in s if str(x).strip()])
            agg = summary.groupby("CostHead", as_index=False).agg(agg_columns)
        else:
            agg_columns = {
                "ITEM Remark": lambda s: "; ".join([x for x in s if str(x).strip()])
            }
            if "GST Items" in summary.columns:
                agg_columns["GST Items"] = lambda s: "\n".join([x for x in s if str(x).strip()])
            agg = summary.groupby("CostHead", as_index=False).agg(agg_columns)
        summary = agg

    # Pad missing heads and enforce order
    order_df = pd.DataFrame({"CostHead": desired_order})
    summary = order_df.merge(summary, on="CostHead", how="left")
    if "Material TotalAmount" in summary.columns:
        summary["Material TotalAmount"] = summary["Material TotalAmount"].fillna("")
    if "GST Amount" in summary.columns:
        summary["GST Amount"] = summary["GST Amount"].fillna(0.0)
    if "Total Material Amount with GST" in summary.columns:
        summary["Total Material Amount with GST"] = summary["Total Material Amount with GST"].fillna("")
    summary["ITEM Remark"] = summary["ITEM Remark"].fillna("")
    if "GST Items" in summary.columns:
        summary["GST Items"] = summary["GST Items"].fillna("")

    # After reordering/padding, compute the Indoor Amenities Excel row index at write-time
    indoor_item_excel_row = None

##TODO: For Source Add This
    # def make_sources(g: pd.DataFrame) -> str:
    #     tuples = (
    #         g[["ActivityName","ParentWBS","SubProject","Project"]]
    #         .drop_duplicates().astype(str)
    #         .agg(" | ".join, axis=1).tolist()
    #     )
    #     return "; ".join(tuples)[:60000]

    # srcs = (
    #     matched_final.groupby("CostHead")[
    #         ["ActivityName","ParentWBS","SubProject","Project"]
    #     ]
    #     .apply(make_sources)
    #     .reset_index(name="Sources")
    # )
    # summary = summary.merge(srcs, on="CostHead", how="left")

    # For audit, show what *would have been* unmatched BEFORE we reallocated (use snapshot)
    other_before = tagged_before_fallbacks[tagged_before_fallbacks["CostHead"] == "Other"].copy()

    # Amenities: rows without ActivityName or ParentWBS
    def _fmt_amenity_name(sp: str) -> str:
        s = _norm_sub(sp)
        if s.startswith("TOWER ") and len(s) == len("TOWER X"):
            return s.replace(" ", " - ") + " ( No Activity Code )"
        if s == "PODIUM" or s == "NON TOWER AREA":
            return "PODIUM / NON TOWER AREA ( No Activity Code )"
        return f"{s} ( No Activity Code )" if s else "( No Activity Code )"

    amenities_mask = (other_before["ActivityName"].fillna("") == "") | (other_before["ParentWBS"].fillna("") == "")
    amenities_rows = other_before[amenities_mask].copy()
    # Add Amenity column
    amenities_rows = amenities_rows.assign(Amenity=amenities_rows["SubProject"].apply(_fmt_amenity_name))

    def _token(row):
        ig = str(row.get("ItemGroup", "") or "").strip()
        idc = str(row.get("ItemDesc", "") or "").strip()
        if ig and idc:
            return f"{ig} | {idc}"
        return idc or ig or "(blank)"

    amenities_items = []
    amenities_gst_items = []
    if not amenities_rows.empty:
        for amenity, group in amenities_rows.groupby("Amenity"):
            items = "; ".join(sorted(set(_token(r) for _, r in group.iterrows())))
            amenities_items.append({"Amenity": amenity, "ITEM Remark": items})

            # Create GST items for amenities
            if has_gst_data:
                gst_items = [item for item in [_gst_item_token(r) for _, r in group.iterrows()] if item]
                gst_items_uniq = sorted(set(gst_items))
                amenities_gst_items.append({"Amenity": amenity, "GST Items": "\n".join(gst_items_uniq) if gst_items_uniq else ""})
            else:
                amenities_gst_items.append({"Amenity": amenity, "GST Items": ""})
    amenities_items = pd.DataFrame(amenities_items)
    amenities_gst_items_df = pd.DataFrame(amenities_gst_items)
    if not amenities_rows.empty:
        if has_gst_data and "GST Amount" in amenities_rows.columns:
            amenities_summary = (
                amenities_rows.groupby("Amenity", as_index=False)
                .agg({
                    "Amount": "sum",
                    "GST Amount": "sum"
                })
                .rename(columns={"Amount": "Material TotalAmount"})
            )
            if not amenities_items.empty:
                amenities_summary = amenities_summary.merge(amenities_items, on="Amenity", how="left")
            if not amenities_gst_items_df.empty:
                amenities_summary = amenities_summary.merge(amenities_gst_items_df, on="Amenity", how="left")
            amenities_summary["Total Material Amount with GST"] = amenities_summary["Material TotalAmount"] + amenities_summary["GST Amount"]
        else:
            amenities_summary = (
                amenities_rows.groupby("Amenity", as_index=False)["Amount"]
                .sum().rename(columns={"Amount":"Material TotalAmount"})
            )
            if not amenities_items.empty:
                amenities_summary = amenities_summary.merge(amenities_items, on="Amenity", how="left")
            if not amenities_gst_items_df.empty:
                amenities_summary = amenities_summary.merge(amenities_gst_items_df, on="Amenity", how="left")
            amenities_summary["GST Amount"] = 0.0
            amenities_summary["Total Material Amount with GST"] = amenities_summary["Material TotalAmount"]
    else:
        amenities_summary = pd.DataFrame(columns=["Amenity", "Material TotalAmount", "GST Amount", "Total Material Amount with GST", "ITEM Remark", "GST Items"])

    # Include amenities also in unmatched views
    # IMPORTANT: Build unmatched sheets from FINAL allocation so anything
    # that got allocated by fallbacks/mapping does NOT remain in unmatched
    final_other = tagged_final[tagged_final["CostHead"] == "Other"].copy()
    unmatched_detail = (
        final_other[["ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc","IssueQty","Amount"]]
        .copy().rename(columns={"Amount":"IssueAmt"})
        .sort_values(["SubProject","ActivityName","ParentWBS","Project"])
    )
    unmatched_by_sp = (
        final_other.groupby(["SubProject","ActivityName","ParentWBS","Project"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["SubProject","TotalAmount"], ascending=[True, False])
    )

    # Append amenities summary at end of CostHead_Summary
    amenities_as_costhead = amenities_summary.rename(columns={"Amenity":"CostHead"})[
        ["CostHead","Material TotalAmount","GST Amount","Total Material Amount with GST","ITEM Remark","GST Items"]
    ]
    amenities_header = pd.DataFrame({
        "CostHead": ["Extra Remaining -Unmatched"],
        "Material TotalAmount": [""],
        "GST Amount": [""],
        "Total Material Amount with GST": [""],
        "ITEM Remark": [""],
        "GST Items": [""]
    })
    # Track the position (Excel row) of the amenities header for formatting
    summary_base_len = len(summary)
    summary = pd.concat([summary, amenities_header, amenities_as_costhead], ignore_index=True)

    # Build special section for excluded unmatched items by keywords for allowed SubProjects only
    def _is_allowed_sp_for_unmatched(sp: str) -> bool:
        return _is_allowed_subproject(sp)

    # Rows that remained Other before fallbacks AND were excluded by keywords
    excluded_unmatched = other_before.copy()
    if "__ExcludedByKeywords__" in tagged.columns:
        # Merge the flag from current tagged by index alignment
        flag_series = tagged.get("__ExcludedByKeywords__").fillna(False)
        if len(flag_series) == len(excluded_unmatched):
            excluded_unmatched["__ExcludedByKeywords__"] = flag_series.values
        else:
            # Best-effort: recompute
            excluded_unmatched["__ExcludedByKeywords__"] = excluded_unmatched.apply(
                lambda r: _is_excluded_for_unmatched(r.get("ItemGroup"), r.get("ItemDesc"), r.get("Remarks")), axis=1
            )
    else:
        excluded_unmatched["__ExcludedByKeywords__"] = excluded_unmatched.apply(
            lambda r: _is_excluded_for_unmatched(r.get("ItemGroup"), r.get("ItemDesc"), r.get("Remarks")), axis=1
        )

    # Keep only excluded with Activity (per latest rule)
    excluded_unmatched = excluded_unmatched[(excluded_unmatched["__ExcludedByKeywords__"] == True) & (excluded_unmatched["ActivityName"].fillna("") != "")]
    # Keep only Podium or Tower A-Z
    excluded_unmatched = excluded_unmatched[excluded_unmatched["SubProject"].apply(_is_allowed_sp_for_unmatched)]

    if not excluded_unmatched.empty:
        # Group by SubProject and aggregate amounts and items
        ex_items = []
        ex_gst_items = []
        for subproject, group in excluded_unmatched.groupby("SubProject"):
            items = "; ".join(sorted(set(_token(r) for _, r in group.iterrows())))
            ex_items.append({"SubProject": subproject, "ITEM Remark": items})

            # Create GST items for excluded unmatched
            if has_gst_data:
                gst_items = [item for item in [_gst_item_token(r) for _, r in group.iterrows()] if item]
                gst_items_uniq = sorted(set(gst_items))
                ex_gst_items.append({"SubProject": subproject, "GST Items": "\n".join(gst_items_uniq) if gst_items_uniq else ""})
            else:
                ex_gst_items.append({"SubProject": subproject, "GST Items": ""})
        ex_items = pd.DataFrame(ex_items)
        ex_gst_items_df = pd.DataFrame(ex_gst_items)
        if has_gst_data and "GST Amount" in excluded_unmatched.columns:
            ex_summary = (
                excluded_unmatched.groupby("SubProject", as_index=False)
                .agg({
                    "Amount": "sum",
                    "GST Amount": "sum"
                })
                .rename(columns={"Amount": "Material TotalAmount"})
                .sort_values("SubProject")
            )
            ex_summary["Total Material Amount with GST"] = ex_summary["Material TotalAmount"] + ex_summary["GST Amount"]
        else:
            ex_summary = (
                excluded_unmatched.groupby("SubProject", as_index=False)["Amount"]
                .sum().rename(columns={"Amount":"Material TotalAmount"})
                .sort_values("SubProject")
            )
            ex_summary["GST Amount"] = 0.0
            ex_summary["Total Material Amount with GST"] = ex_summary["Material TotalAmount"]
        ex_summary = ex_summary.merge(ex_items, on="SubProject", how="left")
        if not ex_gst_items_df.empty:
            ex_summary = ex_summary.merge(ex_gst_items_df, on="SubProject", how="left")
        unmatched_items_as_costhead = ex_summary.rename(columns={"SubProject":"CostHead"})[
            ["CostHead","Material TotalAmount","GST Amount","Total Material Amount with GST","ITEM Remark","GST Items"]
        ]
        unmatched_items_header = pd.DataFrame({
            "CostHead": ["Unmatched Item Group & Items"],
            "Material TotalAmount": [""],
            "GST Amount": [""],
            "Total Material Amount with GST": [""],
            "ITEM Remark": [""],
            "GST Items": [""]
        })
        # Append after amenities block
        unmatched_base_len = len(summary)
        summary = pd.concat([summary, unmatched_items_header, unmatched_items_as_costhead], ignore_index=True)

    # 4) Create output directory and write reports
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = output_path / "CostHead_Reports.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as xw:
        breakdown.to_excel(xw, sheet_name="CostHead_Breakdown", index=False)
        summary.to_excel(xw, sheet_name="CostHead_Summary", index=False)
        map_expanded.to_excel(xw, sheet_name="Mapping_Expanded", index=False)
        unmatched_detail.to_excel(xw, sheet_name="Unmatched_Detail", index=False)
        unmatched_by_sp.to_excel(xw, sheet_name="Unmatched_By_SubProject", index=False)

        # Apply formatting to the Amenities header row in CostHead_Summary
        try:
            ws = xw.sheets["CostHead_Summary"]
            header_row_idx = summary_base_len + 2  # +1 for header, +1 for 1-based rows
            n_cols = summary.shape[1]
            fill = PatternFill(fill_type="solid", start_color="FFEFEFEF", end_color="FFEFEFEF")
            font = Font(bold=True)
            align = Alignment(horizontal="left")
            for c in range(1, n_cols + 1):
                cell = ws.cell(row=header_row_idx, column=c)
                cell.fill = fill
                cell.font = font
                cell.alignment = align
        except Exception:
            pass

        # Wrap text for Indoor Amenities ITEM cell so bullet list is readable
        try:
            ws = xw.sheets["CostHead_Summary"]
            # Recompute the row of Indoor amenities after ordering/padding
            try:
                indoor_idx_zero = summary.index[summary["CostHead"].astype(str).str.strip().str.lower() == "indoor amenities"][0]
                indoor_item_excel_row = int(indoor_idx_zero) + 2
            except Exception:
                indoor_item_excel_row = None
            if indoor_item_excel_row is not None:
                item_col_idx = summary.columns.get_loc("ITEM Remark") + 1
                cell = ws.cell(row=indoor_item_excel_row, column=item_col_idx)
                cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")
        except Exception:
            pass

        # Apply formatting to the Unmatched Item Group & Items header row similar to amenities header
        try:
            ws = xw.sheets["CostHead_Summary"]
            # Find the row index of the unmatched header by searching first column values
            # Since we appended in order, scan for exact header text
            n_rows = summary.shape[0]
            header_row_idx = None
            for i in range(2, n_rows + 2):  # 1-based rows with header
                val = ws.cell(row=i, column=1).value
                if str(val).strip().upper() == "UNMATCHED ITEM GROUP & ITEMS":
                    header_row_idx = i
                    break
            if header_row_idx is not None:
                n_cols = summary.shape[1]
                fill = PatternFill(fill_type="solid", start_color="FFEFEFEF", end_color="FFEFEFEF")
                font = Font(bold=True)
                align = Alignment(horizontal="left")
                for c in range(1, n_cols + 1):
                    cell = ws.cell(row=header_row_idx, column=c)
                    cell.fill = fill
                    cell.font = font
                    cell.alignment = align
        except Exception:
            pass

    return {
        "output_file": str(output_file),
        "output_dir": str(output_path),
        "sheets": ["CostHead_Breakdown", "CostHead_Summary", "Mapping_Expanded", "Unmatched_Detail", "Unmatched_By_SubProject"]
    }
