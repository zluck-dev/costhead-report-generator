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

# REGULAR HARDWARE + these descriptions must not be Masonry; they are forced to Other and
# appear under "Unmatched Item Group & Items" (see compulsory masonry exclusion).
MASONRY_STRICT_GROUP_EXCLUDES = {
    "REGULAR HARDWARE": [
        "L & KEY LOCK 32MM",
        "SS HINGES 5X12",
        "TABLE CHAIN 12",
        "CLEAR SILICONE 789 DOWSEAL",
        "NEUTRAL SILICON CLEAR",
        "FEVICOL HEATEX",
        "FEVICOL HETAX",
        "TENEX CLEAR",
        "3 4 ALUMINIUM PATTI",
        "T ANGLE",
    ]
}

COMPULSORY_MASONARY_EXCLUDE_DESC = [
    '8" PVC PIPE',
    '1 1 2" NON ISI UPVC BEND 45',
    '1 1 2" NON ISI UPVC UNION',
    '160MM NON ISI PVC PIPE 6KG 6MTR',
    '2" NON ISI CPVC COUPLER',
    '200MM NON ISI PVC PIPE 6KG 6MTR',
    '25MM NON ISI PVC PIPE 20KG 6MTR',
    '50MM NON ISI PVC PIPE 10KG 6MTR',
    '50MM NON ISI PVC PIPE 6KG 6MTR',
    '75MM NON ISI PVC PIPE 10KG 6MTR',
    '75MM NON ISI PVC PIPE 6KG 6MTR',
    'NON ISI UPVC SOLVENT 237ML',
    '6" HANDLE',
    'ABRO TAP 1',
    'ABRO TAPE',
    'ABRO TAPE 2',
    'ABRO TAPE 3 4',
    'ALDROP SS 10',
    'BEARING PIVOT 100MM X 19MM',
    'BOLT 8MM X 1',
    'BOLT 8MM X 3',
    'BROWN TAPE',
    'CLEAR SILICON 300ML',
    'DOOR ROLLER GOLI',
    'FEVIKWIK',
    "FLOOR PROTECTION SHEET SIZE 6'X4' 2.5MM THIKNESS",
    'HEX SCREW 25X12',
    'KHILI 17 X 1',
    'KHILI STEEL 500GM 14*1.50',
    'KHILI STEEL 500GM 17*1',
    'KHILI STEEL 500GM 17*1.25',
    'KHILI STEEL 500GM 19*3 4',
    'L CLAMP NUT BOLT',
    'MS HINGIS 5',
    'PAD LOCK 65MM',
    'PROTECTION STICKER GUM GREEN ROLL 20 INCH X 45 MTR',
    'PTA SCREW 19X6',
    'PUSH MAGNET',
    'PVC WALL PLUG 10*35',
    'PVC WALL PLUG 12*50',
    'SCREW 25 X 7',
    'SCREW 50 X 10',
    'SCREW 75 X 10',
    'SELF DRILLING SCREW PTA 150*10',
    'SELF DRILLING SCREW PTA 19*7',
    'SELF DRILLING SCREW PTA 25*7',
    'SELF DRILLING SCREW PTA 32*8',
    'SELF DRILLING SCREW PTA 38*8',
    'SELF DRILLING SCREW PTA 50*8',
    'SELF DRILLING SCREW PTA 60*10',
    'SELF DRILLING SCREW PTA 60*8',
    'SELF DRILLING SCREW PTA 75*10',
    'SELF DRILLING SCREW PTA 75*8',
    'SS HINGES 5X1.25',
    'SS HINGIS',
    'SS TADI 200MM 8',
    'TEFLONE TAPE 12MM X 10MTR',
    'TILE SPACER 3MM',
    'TILE SPACER 4MM',
    'TUBULAR LOCK',
    'WALL MAGNET',
    'WALL PLUG 12 X 35',
    'WALL PLUG 12 X 50',
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

COMMON_PLUMBING_EXCLUDE_GROUPS = [
    "ALUMINIUM COMPOSITE PANEL",
    "ALUMINIUM SQUARE PIPE",
    "CHINA MOSAIC TILE",
    "CHINA MOSAIC",
    "COMMERCIAL PLY",
    "FLUSH DOOR",
    "LAMINATE",
    "TILE ADHESIVE ITEMS",
    "TILE ADHESIVE",
    "ADHESIVE",
    "WHITE CEMENT",
]


COMPULSORY_COMMONPLUMING_EXCLUDE_DESC = [
    'AL 43 MT 2201 BLACK MATT ACP 4MM 0.25 (OUTSIDE LIVING AREA )',
    'ALUCOWIN PS 13 PEARL BROWN SIZE 8X4 THICKNESS 3MM 0.25'
]


COMPULSORY_TERRACEFINISHING_EXCLUDE_DESC = [
    'AL 43 MT 2201 BLACK MATT ACP 4MM 0.25 FR GRADED(B CLASS)'
]

COMPULSORY_COMMONELECTRIC_EXCLUDE_DESC = [
    '70005698 25MM CPVC COUPLER',
    '70001133 75MM SWR COUPLER PF',
    '70001161 75MM SWR SINGLE TEE PF',
    '70002032 2 S S 3M 10FT 75MM A SWR P F',
    '2" HOSE PIPE'
]

COMPULSORY_CONCRETE_EXCLUDE_DESC = [
    'WHITE MARBLE (REGULAR THICKNESS: 15–16MM)'
]

CONCRETE_STRICT_GROUP_EXCLUDES = {
    "FABRICATION": ["PROVIDING OF POLY CARBONATE SHEET"],
    "M.S. SQUARE PIPE": ["ERW MS PIPES"],
}

LANDSCAPE_STRICT_GROUP_EXCLUDES = {
    "FLUSH TANK": ["GEBERIT FLUSH TANK PLATE"],
    "HDF": ["6MM 8X4 HDF"],
}

COMPULSORY_LIFT_EXCLUDE_DESC = [
    '8MM COMMERCIAL MDF 8*4',
    'COMMERCIAL PLY 18MM 8*4',
    'COMMERCIAL PLY 6MM 8*4',
    'EXHAUST FAN 12"',
    'M.S. PLATE 5MM',
    'NATURAL STONE',
    'WEBERFIX PU HIFLEX (R2T) 5KG BUCKET PACKING (COMPONENT A 2.5KG , COMPONENT B 2.5KG)',
    'SIKACERAM 125 EASYFIX GREY (C1T) 40KG',
    'SIKACERAM 255 GREY (C2TE) 25KG',
    'SIKACERAM 288H WHITE (C2TES1) 25KG',
    'SIKACERAM 255 WHITE (C2TE) 25KG',
    'SIKACERAM 288H GREY (C2TES1) 25KG',
    'WEBERFIX PU ( 10KG )',
    'WEBERSET FIRM GREY (C2TE) 20 KG',
    'COLD BOND 1 LTD',
    'OLD VALSADI WOOD'
]

COMPULSORY_WINDOWSECTION_KEYWORDS_DESC = [
    'SIKACERAM 125 EASYFIX GREY (C1T) 40KG'
]

# Keywords for RAILING, GRILL classification
RAILING_GRILL_KEYWORDS_DESC = [

]
RAILING_GRILL_KEYWORDS_GROUP = [
    "BALCONY RAILING",
]
RAILING_GRILL_EXCLUDE_DESC: List[str] = [
    # Add desc-level exclusions here if needed, e.g. "DOOR FRAME"
]
RAILING_GRILL_EXCLUDE_GROUP: List[str] = [
    # Add group-level exclusions here if needed
]

# Keywords for Louvers classification
LOUVERS_KEYWORDS_DESC = [

]
LOUVERS_KEYWORDS_GROUP = [
    "ALUMINIUM LOUVERS"
]
LOUVERS_EXCLUDE_DESC: List[str] = [

]
LOUVERS_EXCLUDE_GROUP: List[str] = [

]

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
            # Use word boundaries only if the keyword starts/ends with a word character
            import re
            k_esc = re.escape(keyword_upper)
            start_b = r'\b' if keyword_upper[0].isalnum() else ''
            end_b = r'\b' if keyword_upper[-1].isalnum() else ''
            pattern = rf'{start_b}{k_esc}{end_b}'
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

def _regular_hardware_masonry_desc_excluded(item_group: str, hay: str) -> bool:
    """Normalized ItemDesc+Remarks: excluded from Masonry when ItemGroup is REGULAR HARDWARE."""
    g_norm = _norm_text(item_group)
    if g_norm not in MASONRY_STRICT_GROUP_EXCLUDES:
        return False
    
    hn = _norm_text(hay or "")
    # Strip quotes for word-boundary match (e.g. 12" -> 12, 3/4" -> 3 4)
    hn2 = re.sub(r'["\u201c\u201d]', " ", hn)
    hn2 = re.sub(r"\s+", " ", hn2).strip()

    exclude_list = MASONRY_STRICT_GROUP_EXCLUDES[g_norm]
    if _contains_exact_keyword(hn2, exclude_list):
        return True
    return False

def _is_regular_hardware_masonry_excluded(item_group: str, item_desc: str, remarks: str) -> bool:
    return _regular_hardware_masonry_desc_excluded(item_group, f"{item_desc} {remarks}")

def _is_compulsory_masonry_exclude_for_row(row: Dict) -> bool:
    item_group = str(row.get("ItemGroup", "") or "")
    item_desc = str(row.get("ItemDesc", "") or "")
    remarks = str(row.get("Remarks", "") or "")

    g_norm = _norm_text(item_group)
    # If the group has strict exclusions, ONLY those items are excluded.
    if g_norm in MASONRY_STRICT_GROUP_EXCLUDES:
        return _regular_hardware_masonry_desc_excluded(item_group, f"{item_desc} {remarks}")

    # For other groups, check the broad compulsory exclusion list.
    d = item_desc.upper()
    r = remarks.upper()
    hay = f"{d} {r}".strip()
    return _contains_exact_keyword(hay, COMPULSORY_MASONARY_EXCLUDE_DESC)

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
        # Respect Masonry excludes; selected Regular Hardware lines go to Unmatched, not Masonry
        if not _is_regular_hardware_masonry_excluded(g, d, r) and not (
            _contains_exact_keyword(hay_desc, MASONRY_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, MASONRY_EXCLUDE_GROUP)
        ):
            return "Masonry and plaster material only"

    if sp_allowed:
        # Respect Concrete excludes
        for kw in CONCRETE_KEYWORDS_DESC:
            if _contains_exact_keyword(hay_desc, [kw]) and not (_contains_exact_keyword(hay_desc, CONCRETE_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, CONCRETE_EXCLUDE_GROUP)):
                return "Concrete"
        for kw in CONCRETE_KEYWORDS_GROUP:
            if _contains_exact_keyword(hay_group, [kw]) and not (_contains_exact_keyword(hay_desc, CONCRETE_EXCLUDE_DESC) or _contains_exact_keyword(hay_group, CONCRETE_EXCLUDE_GROUP)):
                return "Concrete"

    # RAILING, GRILL — applies to ALL SubProjects (no sp_allowed restriction)
    railing_hit = (
        _contains_exact_keyword(hay_desc, RAILING_GRILL_KEYWORDS_DESC) or
        _contains_exact_keyword(hay_group, RAILING_GRILL_KEYWORDS_GROUP)
    )
    if railing_hit:
        if not (_contains_exact_keyword(hay_desc, RAILING_GRILL_EXCLUDE_DESC) or
                _contains_exact_keyword(hay_group, RAILING_GRILL_EXCLUDE_GROUP)):
            return "RAILING, GRILL"

    # Louvers — applies to ALL SubProjects (no sp_allowed restriction)
    louvers_hit = (
        _contains_exact_keyword(hay_desc, LOUVERS_KEYWORDS_DESC) or
        _contains_exact_keyword(hay_group, LOUVERS_KEYWORDS_GROUP)
    )
    if louvers_hit:
        if not (_contains_exact_keyword(hay_desc, LOUVERS_EXCLUDE_DESC) or
                _contains_exact_keyword(hay_group, LOUVERS_EXCLUDE_GROUP)):
            return "Louvers"

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
    railing_keys = {"railing grill", "railing, grill", "railing grill", "ms railing grill"}
    if key in railing_keys or "railing" in key and "grill" in key:
        return "RAILING, GRILL"
    if key == "louvers":
        return "Louvers"
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
    unit_finish_head = "UNIT FINISH (COST REQUIRE TO FINISH ONE. UNIT (FLAT SHOP OFFICE SERVANT ROOM) . EXCLUDING MASONRY & PLASTER"
    unit_finish_norm = _norm_text(unit_finish_head)

    # Build expanded mapping rules:
    # - Keep Steel/Concrete in special-keyword classification only
    # - Keep Masonry rules separately so they can run before keyword fallback
    expanded = []
    masonry_expanded = []
    special_heads = {"Steel", "Concrete"}
    # Priority field order: ActivityName matches are applied before broader fields
    # so that specific role-based keywords (e.g. "LIFT") are not accidentally
    # captured by a rule that matches on ParentWBS for a different cost head.
    HIGH_PRIORITY_FIELDS = {"ActivityName"}

    for _, r in mapping.iterrows():
        head_raw = r["CostHead"]
        head_can = _canonicalize_costhead(head_raw)
        fields_raw = str(r["FromWhere"] or "")
        fields = [f.strip() for f in re.split(r"[;,]", fields_raw) if f.strip()]
        if not fields:
            fields = [_field_for(r["FromWhere"])]
        for kw in _split_keywords(r["Value"]):
            if head_raw and kw:
                for f in fields:
                    field = _field_for(f)
                    priority = 1 if field in HIGH_PRIORITY_FIELDS else 2
                    if head_can in special_heads:
                        continue
                    if head_can == "Masonry and plaster material only":
                        masonry_expanded.append({"CostHead": head_raw, "Field": field, "Keyword": kw, "Priority": priority})
                    else:
                        head_order = 0 if _norm_text(head_raw) == unit_finish_norm else 1
                        expanded.append({"CostHead": head_raw, "Field": field, "Keyword": kw, "Priority": priority, "HeadOrder": head_order})
    map_expanded = pd.DataFrame(expanded)
    masonry_map_expanded = pd.DataFrame(masonry_expanded)

    # Apply assignment in explicit stages:
    #   Stage 1) Special keyword heads (Steel/Concrete)
    #   Stage 2) Unit Finish mapping (priority 1 then 2)
    #   Stage 3) Masonry (mapping first, then Masonry keywords)
    #   Stage 4) RAILING, GRILL + Louvers keywords
    #   Stage 5) Remaining mappings (priority 1 then 2)
    # Preserve ActivityName-before-broader-field behavior within mapping stages.
    assigned_by_mapping = 0

    # Stage 1: Apply special keyword heads (Steel/Concrete) first
    specials = tagged.apply(_classify_special_head, axis=1)
    special_mask = (~assigned_mask) & specials.isin(["Steel", "Concrete"])
    if special_mask.any():
        tagged.loc[special_mask, "CostHead"] = specials[special_mask]
        assigned_mask = assigned_mask | special_mask

    if not map_expanded.empty:
        # Stage 2: Unit Finish mapping
        for pass_priority in (1, 2):
            pass_rules = map_expanded[
                (map_expanded["HeadOrder"] == 0) &
                (map_expanded["Priority"] == pass_priority)
            ]
            for _, mr in pass_rules.iterrows():
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
                    assigned_by_mapping += int(to_assign.sum())

    # Stage 3A: Apply Masonry mapping rules to still-unassigned rows
    if not masonry_map_expanded.empty:
        for pass_priority in (1, 2):
            pass_rules = masonry_map_expanded[
                (masonry_map_expanded["Priority"] == pass_priority)
            ]
            for _, mr in pass_rules.iterrows():
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
                    assigned_by_mapping += int(to_assign.sum())

    # Stage 3B: Masonry keyword fallback on still-unassigned rows
    specials = tagged.apply(_classify_special_head, axis=1)
    masonry_special_mask = (~assigned_mask) & specials.eq("Masonry and plaster material only")
    if masonry_special_mask.any():
        tagged.loc[masonry_special_mask, "CostHead"] = specials[masonry_special_mask]
        assigned_mask = assigned_mask | masonry_special_mask

    # Stage 4: RAILING, GRILL + Louvers keyword fallback on still-unassigned rows
    railing_louvers_mask = (~assigned_mask) & specials.isin(["RAILING, GRILL", "Louvers"])
    if railing_louvers_mask.any():
        tagged.loc[railing_louvers_mask, "CostHead"] = specials[railing_louvers_mask]
        assigned_mask = assigned_mask | railing_louvers_mask

    # Stage 5: Apply all remaining mapping heads
    if not map_expanded.empty:
        for pass_priority in (1, 2):
            pass_rules = map_expanded[
                (map_expanded["HeadOrder"] == 1) &
                (map_expanded["Priority"] == pass_priority)
            ]
            for _, mr in pass_rules.iterrows():
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
                    assigned_by_mapping += int(to_assign.sum())

    # Drop the Priority helper column before returning so callers aren't affected
    if "Priority" in map_expanded.columns:
        map_expanded = map_expanded.drop(columns=["Priority"])
    if "HeadOrder" in map_expanded.columns:
        map_expanded = map_expanded.drop(columns=["HeadOrder"])

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

    # 2.5) Compulsory Masonry Exclusion (Late Override)
    # This ensures items in COMPULSORY_MASONARY_EXCLUDE_DESC are never in Masonry Coshead
    masonry_norm = _norm_text("Masonry and plaster material only")

    # Check for rows that are currently Masonry but should be excluded
    is_masonry_mask = tagged_final["CostHead"].apply(_norm_text) == masonry_norm
    comp_exclude_mask = tagged_final.apply(_is_compulsory_masonry_exclude_for_row, axis=1)

    final_exclude_mask = is_masonry_mask & comp_exclude_mask
    tagged_final["__CompulsoryMasonryExcluded__"] = False
    if final_exclude_mask.any():
        tagged_final.loc[final_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_exclude_mask, "__CompulsoryMasonryExcluded__"] = True

    # 2.6) Compulsory Common Plumbing Exclusion (Late Override)
    # This ensures items in COMPULSORY_COMMONPLUMING_EXCLUDE_DESC are never in Common Plumbing Coshead
    common_plumbing_norm = _norm_text("COMMON PLUMBING (INCLUDING PUMPS)")

    def _is_compulsory_commonplumbing_exclude(row):
        item_group = _norm_text(str(row.get("ItemGroup", "") or ""))
        if item_group in COMMON_PLUMBING_EXCLUDE_GROUPS:
            return True

        d = str(row.get("ItemDesc", "") or "").upper()
        # Common plumbing exclusion must match ItemDesc only (no Remarks).
        hay = d.strip()
        return _contains_exact_keyword(hay, COMPULSORY_COMMONPLUMING_EXCLUDE_DESC)

    # Check for rows that are currently Common Plumbing but should be excluded
    is_common_plumbing_mask = tagged_final["CostHead"].apply(_norm_text) == common_plumbing_norm
    common_comp_exclude_mask = tagged_final.apply(_is_compulsory_commonplumbing_exclude, axis=1)

    final_common_exclude_mask = is_common_plumbing_mask & common_comp_exclude_mask
    tagged_final["__CompulsoryCommonPlumbingExcluded__"] = False
    if final_common_exclude_mask.any():
        tagged_final.loc[final_common_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_common_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_common_exclude_mask, "__CompulsoryCommonPlumbingExcluded__"] = True

    # 2.7) Compulsory Terrace Finishing Exclusion (Late Override)
    # This ensures items in COMPULSORY_TERRACEFINISHING_EXCLUDE_DESC are never in Terrace Finishing Coshead
    terrace_finishing_norm = _norm_text("TERRACE FINISHING (WATERPROOFING AND FINISHING)")

    def _is_compulsory_terracefinishing_exclude(row):
        # Terrace finishing exclusion must match ItemDesc only (no Remarks).
        d = str(row.get("ItemDesc", "") or "").upper()
        hay = d.strip()
        return _contains_exact_keyword(hay, COMPULSORY_TERRACEFINISHING_EXCLUDE_DESC)

    # Check for rows that are currently Terrace Finishing but should be excluded
    is_terrace_finishing_mask = tagged_final["CostHead"].apply(_norm_text) == terrace_finishing_norm
    terrace_comp_exclude_mask = tagged_final.apply(_is_compulsory_terracefinishing_exclude, axis=1)

    final_terrace_exclude_mask = is_terrace_finishing_mask & terrace_comp_exclude_mask
    tagged_final["__CompulsoryTerraceFinishingExcluded__"] = False
    if final_terrace_exclude_mask.any():
        tagged_final.loc[final_terrace_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_terrace_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_terrace_exclude_mask, "__CompulsoryTerraceFinishingExcluded__"] = True

    # 2.8) Compulsory Common Electric Exclusion (Late Override)
    # This ensures items in COMPULSORY_COMMONELECTRIC_EXCLUDE_DESC are never in Common Electric Coshead
    common_electric_norm_raw = _norm_text("Common electric (geb, dg, tc to meter room and meter room to flat mcb)")
    # CostHead label may differ by punctuation/spacing. Normalize to alphanumeric-only
    # before doing equality so minor formatting differences do not break the mask.
    def _norm_costhead_alnum(s: str) -> str:
        t = _norm_text(s)
        return re.sub(r"[^A-Z0-9]+", "", t)

    common_electric_norm = _norm_costhead_alnum(common_electric_norm_raw)

    # Normalize exclusion keywords once (helps if ItemDesc has punctuation/spacing differences).
    # We remove everything except alphanumerics and spaces, then collapse whitespace.
    def _norm_alnum_space(s: str) -> str:
        s2 = _norm_text(s)
        s2 = re.sub(r"[^A-Z0-9]+", " ", s2)
        s2 = re.sub(r"\s+", " ", s2).strip()
        return s2

    common_electric_keywords_norm2 = {_norm_alnum_space(k) for k in COMPULSORY_COMMONELECTRIC_EXCLUDE_DESC}

    def _is_compulsory_commonelectric_exclude(row):
        # Common electric exclusion must match ItemDesc only (no Remarks).
        d2 = _norm_alnum_space(str(row.get("ItemDesc", "") or ""))
        # Use substring match after normalization to tolerate minor extra text.
        return any(kw in d2 for kw in common_electric_keywords_norm2)

    common_electric_comp_exclude_mask = tagged_final.apply(
        _is_compulsory_commonelectric_exclude, axis=1
    )

    # Only exclude rows that are currently classified as "Common electric"
    is_common_electric_mask = tagged_final["CostHead"].apply(_norm_costhead_alnum) == common_electric_norm
    final_common_electric_exclude_mask = is_common_electric_mask & common_electric_comp_exclude_mask

    # DEBUG: Print what CostHeads the keyword-matched items currently belong to
    if common_electric_comp_exclude_mask.any():
        debug_rows = tagged_final[common_electric_comp_exclude_mask][["CostHead", "ItemDesc", "ItemGroup"]].copy()
        print("=== DEBUG: Items matching COMPULSORY_COMMONELECTRIC_EXCLUDE_DESC keywords ===")
        for _, dr in debug_rows.iterrows():
            print(f"  CostHead='{dr['CostHead']}' | ItemDesc='{dr['ItemDesc']}' | ItemGroup='{dr['ItemGroup']}'")
        print(f"  Total keyword matches: {common_electric_comp_exclude_mask.sum()}")
        print(f"  Total in Common Electric CostHead: {is_common_electric_mask.sum()}")
        print(f"  Total actually excluded: {final_common_electric_exclude_mask.sum()}")
        print("=== END DEBUG ===")
    tagged_final["__CompulsoryCommonElectricExcluded__"] = False
    if final_common_electric_exclude_mask.any():
        tagged_final.loc[final_common_electric_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_common_electric_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_common_electric_exclude_mask, "__CompulsoryCommonElectricExcluded__"] = True

    # 2.9) Compulsory Concrete Exclusion (Late Override)
    # Only exclude the exact items listed in COMPULSORY_CONCRETE_EXCLUDE_DESC,
    # and ONLY when they are currently classified as Concrete.
    concrete_norm = _norm_text("Concrete")

    def _is_compulsory_concrete_exclude(row):
        item_group = _norm_text(str(row.get("ItemGroup", "") or ""))
        item_desc = str(row.get("ItemDesc", "") or "")
        
        # Group-specific strict exclusions
        if item_group in CONCRETE_STRICT_GROUP_EXCLUDES:
            d_norm = _norm_text(item_desc)
            exclude_list = [_norm_text(k) for k in CONCRETE_STRICT_GROUP_EXCLUDES[item_group]]
            return _contains_exact_keyword(d_norm, exclude_list)

        # For other groups, check the broad compulsory exclusion list.
        # Concrete exclusion must match ItemDesc only (no Remarks).
        d_norm = _norm_text(item_desc)
        concrete_kw_norm = [_norm_text(k) for k in COMPULSORY_CONCRETE_EXCLUDE_DESC]
        return _contains_exact_keyword(d_norm, concrete_kw_norm)

    is_concrete_mask = tagged_final["CostHead"].apply(_norm_text) == concrete_norm
    concrete_comp_exclude_mask = tagged_final.apply(_is_compulsory_concrete_exclude, axis=1)

    final_concrete_exclude_mask = is_concrete_mask & concrete_comp_exclude_mask
    tagged_final["__CompulsoryConcreteExcluded__"] = False
    if final_concrete_exclude_mask.any():
        tagged_final.loc[final_concrete_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_concrete_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_concrete_exclude_mask, "__CompulsoryConcreteExcluded__"] = True

    # 2.10) Compulsory Lift Exclusion (Late Override)
    # This ensures items in COMPULSORY_LIFT_EXCLUDE_DESC are never in Lift CostHead
    lift_norm = _norm_text("Lift")

    def _is_compulsory_lift_exclude(row):
        # Lift exclusion must match ItemDesc only (no Remarks).
        d = str(row.get("ItemDesc", "") or "").upper()
        hay = d.strip()
        return _contains_exact_keyword(hay, COMPULSORY_LIFT_EXCLUDE_DESC)

    # Check for rows that are currently Lift but should be excluded
    is_lift_mask = tagged_final["CostHead"].apply(_norm_text) == lift_norm
    lift_comp_exclude_mask = tagged_final.apply(_is_compulsory_lift_exclude, axis=1)

    final_lift_exclude_mask = is_lift_mask & lift_comp_exclude_mask
    tagged_final["__CompulsoryLiftExcluded__"] = False
    if final_lift_exclude_mask.any():
        tagged_final.loc[final_lift_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_lift_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_lift_exclude_mask, "__CompulsoryLiftExcluded__"] = True

    # 2.11) Compulsory Window Section Exclusion (Late Override)
    # This ensures items in COMPULSORY_WINDOWSECTION_KEYWORDS_DESC are never in Window section CostHead
    window_section_norm = _norm_text("Window section")

    def _is_compulsory_windowsection_exclude(row):
        # Window section exclusion must match ItemDesc only (no Remarks).
        d = str(row.get("ItemDesc", "") or "").upper()
        hay = d.strip()
        return _contains_exact_keyword(hay, COMPULSORY_WINDOWSECTION_KEYWORDS_DESC)

    # Check for rows that are currently Window section but should be excluded
    is_window_section_mask = tagged_final["CostHead"].apply(_norm_text) == window_section_norm
    window_comp_exclude_mask = tagged_final.apply(_is_compulsory_windowsection_exclude, axis=1)

    final_window_exclude_mask = is_window_section_mask & window_comp_exclude_mask
    tagged_final["__CompulsoryWindowSectionExcluded__"] = False
    if final_window_exclude_mask.any():
        tagged_final.loc[final_window_exclude_mask, "CostHead"] = "Other"
        tagged_final.loc[final_window_exclude_mask, "__ExcludedByKeywords__"] = True
        tagged_final.loc[final_window_exclude_mask, "__CompulsoryWindowSectionExcluded__"] = True

    # 2.12) Late Flagging for Unmatched Section
    # This is a final catch-all pass to ensure every item in our exclusion groups
    # is correctly moved to the "Unmatched" section, no matter what it was labeled before.
    
    def _final_strict_exclude_check(row):
        g_norm = _norm_text(str(row.get("ItemGroup", "") or ""))
        item_desc = str(row.get("ItemDesc", "") or "")
        remarks = str(row.get("Remarks", "") or "")
        hay = f"{item_desc} {remarks}"
        d_upper = item_desc.upper().strip()
        hay_upper = hay.upper().strip()
        
        # 1. Check all Global Compulsory Lists
        if _contains_exact_keyword(hay_upper, COMPULSORY_MASONARY_EXCLUDE_DESC): return True
        if _contains_exact_keyword(d_upper, COMPULSORY_COMMONPLUMING_EXCLUDE_DESC): return True
        if _contains_exact_keyword(d_upper, COMPULSORY_TERRACEFINISHING_EXCLUDE_DESC): return True
        if _contains_exact_keyword(d_upper, COMPULSORY_COMMONELECTRIC_EXCLUDE_DESC): return True
        if _contains_exact_keyword(d_upper, COMPULSORY_LIFT_EXCLUDE_DESC): return True
        if _contains_exact_keyword(d_upper, COMPULSORY_WINDOWSECTION_KEYWORDS_DESC): return True
        
        # 2. Masonry/Hardware (Strict Item List)
        if g_norm in MASONRY_STRICT_GROUP_EXCLUDES:
            if _regular_hardware_masonry_desc_excluded(g_norm, hay):
                return True
        
        # 3. Concrete/Fabrication/Pipes (Strict Item List)
        if g_norm in CONCRETE_STRICT_GROUP_EXCLUDES:
            d_norm = _norm_text(item_desc)
            exclude_list = [_norm_text(k) for k in CONCRETE_STRICT_GROUP_EXCLUDES[g_norm]]
            if _contains_exact_keyword(d_norm, exclude_list):
                return True
        
        # 4. Common Plumbing Groups (Whole group excluded)
        if g_norm in COMMON_PLUMBING_EXCLUDE_GROUPS:
            return True
        
        # 5. Landscape/Flush Tank (Strict Item List)
        if g_norm in LANDSCAPE_STRICT_GROUP_EXCLUDES:
            d_norm = _norm_text(item_desc)
            exclude_list = [_norm_text(k) for k in LANDSCAPE_STRICT_GROUP_EXCLUDES[g_norm]]
            if _contains_exact_keyword(d_norm, exclude_list):
                return True
            
        # 5. Global Concrete Keywords
        d_norm = _norm_text(item_desc)
        concrete_kw_norm = [_norm_text(k) for k in COMPULSORY_CONCRETE_EXCLUDE_DESC]
        if _contains_exact_keyword(d_norm, concrete_kw_norm):
            return True

        return False

    # Final catch-all pass to ensure every item in our exclusion groups
    # is correctly moved to the "Other" section and flagged for the "Unmatched" section.
    is_late_exclude = tagged_final.apply(_final_strict_exclude_check, axis=1)
    if is_late_exclude.any():
        tagged_final.loc[is_late_exclude, "CostHead"] = "Other"
        tagged_final.loc[is_late_exclude, "__ExcludedByKeywords__"] = True

    # 2.13) Targeted Reallocation: Painting Material | SADA PRIMER JOTUN MAKE
    # Move this specific item from Lift (or anywhere) to Outer paint and texture.
    painting_realloc_mask = (
        (tagged_final["ItemGroup"].fillna("").apply(_norm_text) == "PAINTING MATERIAL") &
        (tagged_final["ItemDesc"].fillna("").apply(_norm_text) == "SADA PRIMER JOTUN MAKE")
    )
    if painting_realloc_mask.any():
        # Force the move and ensure it's visible in the target CostHead
        tagged_final.loc[painting_realloc_mask, "CostHead"] = "Outer paint and texture"
        # Wipe "LIFT" association from Activity and WBS columns so it doesn't show "LIFT" there
        tagged_final.loc[painting_realloc_mask, "ActivityName"] = "Outer paint and texture"
        tagged_final.loc[painting_realloc_mask, "ParentWBS"] = "Outer paint and texture"
        
        tagged_final.loc[painting_realloc_mask, "__ExcludedByKeywords__"] = False
        # Double check: ensure no Lift-specific exclusion flags remain
        if "__CompulsoryLiftExcluded__" in tagged_final.columns:
            tagged_final.loc[painting_realloc_mask, "__CompulsoryLiftExcluded__"] = False

    # Canonicalize CostHead labels to avoid duplicates like STEEL vs Steel
    tagged_final["CostHead"] = tagged_final["CostHead"].apply(_canonicalize_costhead)

    # 3) Build reports from the final classification
    # Exclude manually excluded items and incomplete amenities from the main reports
    if "__ExcludedByKeywords__" in tagged_final.columns:
        is_excl = tagged_final["__ExcludedByKeywords__"].fillna(False).astype(bool)
    else:
        is_excl = pd.Series(False, index=tagged_final.index)
        
    is_other_head_raw = tagged_final["CostHead"].fillna("").str.strip().str.upper() == "OTHER"
    is_inc_amenity = is_other_head_raw & (
        (tagged_final["ActivityName"].fillna("") == "") | (tagged_final["ParentWBS"].fillna("") == "")
    )
    
    report_mask = is_excl | is_inc_amenity
    
    breakdown = (
        tagged_final[~report_mask].groupby(["CostHead","ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["CostHead","ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc"])
    )

    matched_final = tagged_final.copy()  # everything, since we've allocated
    # Exclude any rows that are flagged for the Unmatched section or are incomplete Amenities
    if "__ExcludedByKeywords__" in matched_final.columns:
        is_excluded_by_kw = matched_final["__ExcludedByKeywords__"].fillna(False).astype(bool)
    else:
        is_excluded_by_kw = pd.Series(False, index=matched_final.index)

    # Case-insensitive check for "Other" CostHead
    is_other_head = matched_final["CostHead"].fillna("").str.strip().str.upper() == "OTHER"
    
    is_incomplete_amenity = is_other_head & (
        (matched_final["ActivityName"].fillna("") == "") | (matched_final["ParentWBS"].fillna("") == "")
    )
    
    # Final mask: Hide if it's a manual exclusion OR an incomplete amenity
    amenities_mask_summary = is_excluded_by_kw | is_incomplete_amenity
    # Treat GST mode as active whenever GST columns exist in GIN mapped data.
    # This allows GST detail text even when GST amounts are 0 for some/all rows.
    has_gst_data = ("GST Amount" in matched_final.columns and "GST Slab" in matched_final.columns)

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
        # No GST data - use original behavior without GST columns
        summary = (
            matched_final[~amenities_mask_summary]
            .groupby("CostHead", as_index=False)["Amount"]
            .sum().rename(columns={"Amount":"Material TotalAmount"})
            .sort_values("CostHead")
        )

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

    def _is_indoor_amenities_head(head) -> bool:
        return str(head or "").strip().lower() == "indoor amenities"

    # Process all cost heads for GST items; ITEM Remark for every head like Steel/Concrete/Masonry
    # (aggregated ItemGroup | ItemDesc tokens), except Indoor Amenities — that row uses a separate
    # SubProject bullet format below.
    # Use the filtered data for building the Item Remark text
    matched_final_filtered = matched_final[~amenities_mask_summary]

    for head in summary["CostHead"].tolist():
        subset = matched_final_filtered[matched_final_filtered["CostHead"] == head]

        if not _is_indoor_amenities_head(head):
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

    # Only add GST Items column if there's actual GST data
    if has_gst_data:
        summary["GST Items"] = summary["CostHead"].map(gst_items_text).fillna("")

    # Build Indoor Amenities breakdown into ITEM cell as bullet list of SubProject totals
    indoor_item_excel_row = None
    indoor_gst_amount_text = ""
    indoor_total_with_gst_text = ""
    try:
        indoor_mask_data = matched_final["CostHead"].astype(str).str.strip().str.lower() == "indoor amenities"
        if indoor_mask_data.any():
            sp_totals = (
                matched_final[indoor_mask_data]
                .groupby("SubProject", as_index=False)["Amount"].sum()
                .sort_values("SubProject")
            )
            # Clean names and build bullet lines for ITEM Remark
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

            if has_gst_data and "GST Amount" in matched_final.columns:
                indoor_indoor = matched_final[indoor_mask_data].copy()
                indoor_indoor["__GST__"] = pd.to_numeric(indoor_indoor["GST Amount"], errors="coerce").fillna(0.0)
                indoor_indoor["__AMT__"] = pd.to_numeric(indoor_indoor["Amount"], errors="coerce").fillna(0.0)
                indoor_indoor["__TOTAL__"] = indoor_indoor["__AMT__"] + indoor_indoor["__GST__"]

                # GST Amount breakdown per SubProject
                gst_sp_totals = (
                    indoor_indoor.groupby("SubProject", as_index=False)["__GST__"]
                    .sum().sort_values("SubProject")
                )
                gst_lines = []
                for _, r in gst_sp_totals.iterrows():
                    sp = str(r["SubProject"]).strip()
                    amt = float(r["__GST__"]) if pd.notnull(r["__GST__"]) else 0.0
                    if sp:
                        gst_lines.append(f"- {sp} - {amt:,.2f} Rs.")
                gst_total = gst_sp_totals["__GST__"].sum() if not gst_sp_totals.empty else 0.0
                gst_lines.append(f"Total - {gst_total:,.2f} Rs.")
                indoor_gst_amount_text = "\n".join(gst_lines)

                # Total Material Amount with GST breakdown per SubProject
                total_sp_totals = (
                    indoor_indoor.groupby("SubProject", as_index=False)["__TOTAL__"]
                    .sum().sort_values("SubProject")
                )
                total_lines = []
                for _, r in total_sp_totals.iterrows():
                    sp = str(r["SubProject"]).strip()
                    amt = float(r["__TOTAL__"]) if pd.notnull(r["__TOTAL__"]) else 0.0
                    if sp:
                        total_lines.append(f"- {sp} - {amt:,.2f} Rs.")
                total_grand = total_sp_totals["__TOTAL__"].sum() if not total_sp_totals.empty else 0.0
                total_lines.append(f"Total - {total_grand:,.2f} Rs.")
                indoor_total_with_gst_text = "\n".join(total_lines)
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
            if has_gst_data:
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
            if has_gst_data and "GST Items" in summary.columns:
                agg_columns["GST Items"] = lambda s: "\n".join([x for x in s if str(x).strip()])
            agg = summary.groupby("CostHead", as_index=False).agg(agg_columns)
        summary = agg

    # Pad missing heads and enforce order
    order_df = pd.DataFrame({"CostHead": desired_order})
    summary = order_df.merge(summary, on="CostHead", how="left")
    
    # Ensure amount columns are numeric and fill with NaN (empty in Excel) not empty strings
    if "Material TotalAmount" in summary.columns:
        summary["Material TotalAmount"] = pd.to_numeric(summary["Material TotalAmount"], errors="coerce")
    if "GST Amount" in summary.columns:
        summary["GST Amount"] = pd.to_numeric(summary["GST Amount"], errors="coerce")
    if "Total Material Amount with GST" in summary.columns:
        summary["Total Material Amount with GST"] = pd.to_numeric(summary["Total Material Amount with GST"], errors="coerce")

    summary["ITEM Remark"] = summary["ITEM Remark"].fillna("")
    if has_gst_data and "GST Items" in summary.columns:
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

    # Helpers for the unmatched logic below
    def _is_allowed_sp_for_unmatched(sp: str) -> bool:
        return _is_allowed_subproject(sp)

    def _token(row):
        ig = str(row.get("ItemGroup", "") or "").strip()
        idc = str(row.get("ItemDesc", "") or "").strip()
        if ig and idc:
            return f"{ig} | {idc}"
        return idc or ig or "(blank)"

    # --- BUILD FINAL SUMMARY SECTIONS ---
    # We use tagged_final to find all items that ended up as "Other"
    final_other = tagged_final[tagged_final["CostHead"] == "Other"].copy()

    # 1. Compulsory Masonry Excluded
    cm_mask = final_other["__CompulsoryMasonryExcluded__"] == True
    compulsory_masonry_rows = final_other[cm_mask].copy()

    # 1b. Compulsory Common Plumbing Excluded
    cp_mask = final_other["__CompulsoryCommonPlumbingExcluded__"] == True
    compulsory_common_plumbing_rows = final_other[cp_mask].copy()
    # 1c. Compulsory Terrace Finishing Excluded
    tpf_mask = final_other["__CompulsoryTerraceFinishingExcluded__"] == True
    compulsory_terrace_finishing_rows = final_other[tpf_mask].copy()

    # 1d. Compulsory Common Electric Excluded
    ce_mask = final_other["__CompulsoryCommonElectricExcluded__"] == True
    compulsory_common_electric_rows = final_other[ce_mask].copy()

    # 1e. Compulsory Concrete Excluded
    cc_mask = final_other["__CompulsoryConcreteExcluded__"] == True
    compulsory_concrete_rows = final_other[cc_mask].copy()

    # 1f. Compulsory Lift Excluded
    lf_mask = final_other["__CompulsoryLiftExcluded__"] == True
    compulsory_lift_rows = final_other[lf_mask].copy()

    # 1g. Compulsory Window Section Excluded
    ws_mask = final_other["__CompulsoryWindowSectionExcluded__"] == True
    compulsory_window_section_rows = final_other[ws_mask].copy()

    special_excluded_mask = cm_mask | cp_mask | tpf_mask | ce_mask | cc_mask | lf_mask | ws_mask

    # 2. Amenities (Missing Activity/WBS and not already in CM)
    # If a row is already excluded by keywords, it should not be treated as
    # "Amenities (Extra Remaining -Unmatched)" even if Activity/WBS is blank.
    amenities_mask = (
        (~special_excluded_mask)
        & ((final_other["ActivityName"].fillna("") == "") | (final_other["ParentWBS"].fillna("") == ""))
        & (final_other["__ExcludedByKeywords__"] != True)
    )
    amenities_rows = final_other[amenities_mask].copy()

    # 3. Unmatched Keyword Excluded (Flagged as either normal or masonry exclusion)
    kw_mask = (~amenities_mask) & (final_other["__ExcludedByKeywords__"] == True)
    excluded_unmatched = final_other[kw_mask].copy()

    # --- Section: Extra Remaining -Unmatched (Amenities) ---
    amenities_items = []
    amenities_gst_items = []
    if not amenities_rows.empty:
        # Use existing formatting for amenity names
        amenities_rows["Amenity"] = amenities_rows["SubProject"].apply(_fmt_amenity_name)
        for amenity, group in amenities_rows.groupby("Amenity"):
            items = "; ".join(sorted(set(_token(r) for _, r in group.iterrows())))
            amenities_items.append({"Amenity": amenity, "ITEM Remark": items})
            if has_gst_data:
                gst_items = [item for item in [_gst_item_token(r) for _, r in group.iterrows()] if item]
                gst_items_uniq = sorted(set(gst_items))
                amenities_gst_items.append({"Amenity": amenity, "GST Items": "\n".join(gst_items_uniq) if gst_items_uniq else ""})
            else:
                amenities_gst_items.append({"Amenity": amenity, "GST Items": ""})

        amenities_summary_df = pd.DataFrame(amenities_items)
        if has_gst_data and "GST Amount" in amenities_rows.columns:
            amenity_agg = (
                amenities_rows.groupby("Amenity", as_index=False)
                .agg({"Amount": "sum", "GST Amount": "sum"})
                .rename(columns={"Amount": "Material TotalAmount"})
            )
            amenity_agg["Total Material Amount with GST"] = amenity_agg["Material TotalAmount"] + amenity_agg["GST Amount"]
        else:
            amenity_agg = (
                amenities_rows.groupby("Amenity", as_index=False)["Amount"]
                .sum().rename(columns={"Amount":"Material TotalAmount"})
            )

        amenity_final_summary = amenity_agg.merge(amenities_summary_df, on="Amenity", how="left")
        if has_gst_data:
            amenity_final_summary = amenity_final_summary.merge(pd.DataFrame(amenities_gst_items), on="Amenity", how="left")

        # Add to summary
        am_header = pd.DataFrame({"CostHead": ["Extra Remaining -Unmatched"]})
        am_data = amenity_final_summary.rename(columns={"Amenity": "CostHead"})
        summary_base_len = len(summary) # track for formatting
        summary = pd.concat([summary, am_header, am_data], ignore_index=True)

    # --- Section: Unmatched Item Group & Items ---
    if not excluded_unmatched.empty:
        # Filter by allowed subprojects
        ex_filtered = excluded_unmatched[excluded_unmatched["SubProject"].apply(_is_allowed_sp_for_unmatched)]
        if not ex_filtered.empty:
            ex_items = []
            ex_gst_items = []
            for subproject, group in ex_filtered.groupby("SubProject"):
                items = "; ".join(sorted(set(_token(r) for _, r in group.iterrows())))
                ex_items.append({"SubProject": subproject, "ITEM Remark": items})
                if has_gst_data:
                    gst_items = [item for item in [_gst_item_token(r) for _, r in group.iterrows()] if item]
                    gst_items_uniq = sorted(set(gst_items))
                    ex_gst_items.append({"SubProject": subproject, "GST Items": "\n".join(gst_items_uniq) if gst_items_uniq else ""})
                else:
                    ex_gst_items.append({"SubProject": subproject, "GST Items": ""})

            ex_agg = (
                ex_filtered.groupby("SubProject", as_index=False)
                .agg({"Amount": "sum", "GST Amount": "sum"} if has_gst_data else {"Amount": "sum"})
                .rename(columns={"Amount": "Material TotalAmount"})
                .sort_values("SubProject")
            )
            if has_gst_data:
                ex_agg["Total Material Amount with GST"] = ex_agg["Material TotalAmount"] + ex_agg["GST Amount"]

            ex_final_summary = ex_agg.merge(pd.DataFrame(ex_items), on="SubProject", how="left")
            if has_gst_data:
                ex_final_summary = ex_final_summary.merge(pd.DataFrame(ex_gst_items), on="SubProject", how="left")

            kw_header = pd.DataFrame({"CostHead": ["Unmatched Item Group & Items"]})
            kw_data = ex_final_summary.rename(columns={"SubProject": "CostHead"})
            summary = pd.concat([summary, kw_header, kw_data], ignore_index=True)

    # Append grand total row at the end of CostHead_Summary
    total_columns = ["Material TotalAmount", "GST Amount", "Total Material Amount with GST"]
    total_row = {"CostHead": "Total Cost"}
    for col in total_columns:
        if col in summary.columns:
            total_row[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0.0).sum()
    if "ITEM Remark" in summary.columns:
        total_row["ITEM Remark"] = ""
    if "GST Items" in summary.columns:
        total_row["GST Items"] = ""
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

    # Write indoor detailed text after totals are computed, so grand totals remain numeric.
    if has_gst_data:
        indoor_mask_summary = summary["CostHead"].astype(str).str.strip().str.lower() == "indoor amenities"
        if indoor_mask_summary.any():
            if indoor_gst_amount_text and "GST Amount" in summary.columns:
                summary.loc[indoor_mask_summary, "GST Amount"] = indoor_gst_amount_text
            if indoor_total_with_gst_text and "Total Material Amount with GST" in summary.columns:
                summary.loc[indoor_mask_summary, "Total Material Amount with GST"] = indoor_total_with_gst_text

    # Re-build unmatched detail views for the dedicated sheets
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

        # Indian numbering system formatting (e.g., 12652545 -> 1,26,52,545)
        # Note: this changes display only (Excel formatting), while calculations remain numeric.
        try:
            INDIAN_NUMBER_FORMAT = '[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00'

            def _apply_number_format(ws, header_row_idx: int, target_col_name: str):
                # Find the column index by matching the header cell text.
                header_cells = [
                    ws.cell(row=header_row_idx, column=c).value
                    for c in range(1, ws.max_column + 1)
                ]
                target_col_idx = None
                for idx, hdr in enumerate(header_cells, start=1):
                    if hdr is None:
                        continue
                    if str(hdr).strip() == target_col_name:
                        target_col_idx = idx
                        break
                if target_col_idx is None:
                    return

                for r in range(header_row_idx + 1, ws.max_row + 1):
                    cell = ws.cell(row=r, column=target_col_idx)
                    cur_val = cell.value
                    if cur_val is None or cur_val == "":
                        continue

                    # Try to convert current value back to float if it was stored as text
                    try:
                        # Clean common numeric noise: commas, non-breaking spaces, etc.
                        import re
                        s_val = str(cur_val).replace(",", "").replace("\u00a0"," ").strip()
                        # Keep only the numeric parts
                        m = re.search(r'[-+]?\d*\.?\d+', s_val)
                        if m:
                            num_val = float(m.group())
                            cell.value = num_val
                            cell.number_format = INDIAN_NUMBER_FORMAT
                    except Exception:
                        # If it's a multiline string (e.g. Indoor Amenities bullets) or header, keep as text
                        pass

            # Breakdown sheet
            ws_breakdown = xw.sheets["CostHead_Breakdown"]
            _apply_number_format(ws_breakdown, 1, "TotalAmount")

            # Summary sheet
            ws_summary = xw.sheets["CostHead_Summary"]
            # Freeze first column (CostHead) for easier scrolling
            # Assumes headers are in row 1, data starts at row 2.
            ws_summary.freeze_panes = "B2"
            _apply_number_format(ws_summary, 1, "Material TotalAmount")
            if "GST Amount" in summary.columns:
                _apply_number_format(ws_summary, 1, "GST Amount")
            _apply_number_format(ws_summary, 1, "Total Material Amount with GST")

            # Unmatched sheets
            ws_unmatched_detail = xw.sheets["Unmatched_Detail"]
            _apply_number_format(ws_unmatched_detail, 1, "IssueAmt")

            ws_unmatched_by_sp = xw.sheets["Unmatched_By_SubProject"]
            _apply_number_format(ws_unmatched_by_sp, 1, "TotalAmount")
        except Exception:
            pass

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
                if "GST Amount" in summary.columns:
                    gst_col_idx = summary.columns.get_loc("GST Amount") + 1
                    gst_cell = ws.cell(row=indoor_item_excel_row, column=gst_col_idx)
                    gst_cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")
                if "Total Material Amount with GST" in summary.columns:
                    total_col_idx = summary.columns.get_loc("Total Material Amount with GST") + 1
                    total_cell = ws.cell(row=indoor_item_excel_row, column=total_col_idx)
                    total_cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")
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
