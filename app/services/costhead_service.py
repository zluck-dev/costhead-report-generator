import pandas as pd
import re
import warnings
from typing import List, Dict, Tuple
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment

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

STEEL_KEYWORDS_DESC = [
    "TMT","REBAR","REINFORCEMENT","MS ROD","TOR STEEL","BINDING WIRE","STEEL BAR",
    "TMT BAR","DOWEL","CHAIR","REO","DEFORMED BAR","RIBBED BAR"
]
STEEL_KEYWORDS_GROUP = ["STEEL"]
STEEL_EXCLUDE_DESC = ["PPC","OPC","CEMENT","C CHANNEL","CHANNEL","MS C CHANNEL","M.S. C - CHANNEL"]
STEEL_EXCLUDE_GROUP = ["CEMENT","PPC","OPC","CHANNEL"]

CONCRETE_KEYWORDS_DESC = [
    "RMC","READY MIX","READY-MIX","READYMIX","TRANSIT MIX","PUMPED CONCRETE",
    "M20","M25","M30","M35","M40","DESIGN MIX","SITE MIX CONCRETE",
    "FLY ASH CONCRETE","GGBS CONCRETE","WITH FLY ASH","WITHOUT FLY ASH","GGBS","GGBFS",
    "OPC CEMENT","OPC","ORDINARY PORTLAND CEMENT",
    "FLY ASH",
    "SLAG 80-100MM","SLAG 40-80MM","SLAG",
    "6MM MAMRI","MAMRI 6MM",
    "GSB",
    "20MM KAPCHI","KAPCHI 20MM"
]
CONCRETE_KEYWORDS_GROUP = ["CONCRETE","CEMENT CONCRETE","RMC","READY MIX","OPC"]

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
MASONRY_KEYWORDS_GROUP = ["AAC","BRICK","BLOCK","MASONRY","PLASTER","PPC","SAND","PVC PIPE","NON ISI PVC"]

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
    return "ActivityName"

# ---------- SubProject Filter ----------
def _norm_sub(s: str) -> str:
    if s is None: return ""
    s = str(s).upper().replace('-', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def _is_allowed_subproject(sp: str) -> bool:
    spn = _norm_sub(sp)
    if spn == "PODIUM":
        return True
    if spn.startswith("TOWER ") and len(spn) == len("TOWER X"):
        letter = spn.split(" ")[-1]
        return len(letter) == 1 and 'A' <= letter <= 'Z'
    return False

# ---------- Special Heads Classifier ----------
def _is_globally_excluded(item_group: str, item_desc: str, remarks: str) -> bool:
    G, D, R = (str(item_group or "").upper(), str(item_desc or "").upper(), str(remarks or "").upper())
    for kw in EXCLUDE_ANY_DESC:
        if kw in D or kw in R:
            return True
    for kw in EXCLUDE_ANY_GROUP:
        if kw in G:
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

    steel_hit = any(kw in hay_desc for kw in STEEL_KEYWORDS_DESC) or any(kw in hay_group for kw in STEEL_KEYWORDS_GROUP)
    if steel_hit:
        if sp_allowed and not (any(x in hay_desc for x in STEEL_EXCLUDE_DESC) or any(x in hay_group for x in STEEL_EXCLUDE_GROUP)):
            return "Steel"

    if sp_allowed and (any(kw in hay_desc for kw in MASONRY_KEYWORDS_DESC) or any(kw in hay_group for kw in MASONRY_KEYWORDS_GROUP)):
        return "Masonry and plaster material only"

    if sp_allowed:
        for kw in CONCRETE_KEYWORDS_DESC:
            if kw in hay_desc:
                return "Concrete"
        for kw in CONCRETE_KEYWORDS_GROUP:
            if kw in hay_group:
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


    # Expand mapping rules FIRST (mapping takes precedence over specials except Steel/Concrete/Masonry targets)
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

    # Apply mapping rules before specials
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

    # Now pre-assign special heads only for still-unassigned rows
    specials = tagged.apply(_classify_special_head, axis=1)
    special_mask = specials.isin(["Steel","Concrete","Masonry and plaster material only"]) & (~assigned_mask)
    tagged.loc[special_mask, "CostHead"] = specials[special_mask]
    assigned_mask = assigned_mask | special_mask


    # Return after applying mapping and specials
    return tagged, map_expanded

# ---------- Reallocate unmatched by SubProject ----------
def reallocate_unmatched_by_subproject(tagged: pd.DataFrame) -> pd.DataFrame:
    t = tagged.copy()
    # Disabled per user request: do not introduce extra CostHeads based on SubProject
    return t

# ---------- Fallbacks based on ParentWBS for unmatched ----------
def apply_parentwbs_fallbacks(tagged: pd.DataFrame) -> pd.DataFrame:
    t = tagged.copy()
    is_other = (t["CostHead"] == "Other")
    wbs_norm = t["ParentWBS"].fillna("").astype(str).map(_norm_text)
    mask_masonry = is_other & (
        wbs_norm.str.contains(r"\bMASONARY WORK\b", regex=True) |
        wbs_norm.str.contains(r"\bINTERNAL PLASTER WORK\b", regex=True)
    )
    mask_concrete = is_other & wbs_norm.str.contains(r"\bRCC WORK\b", regex=True)
    t.loc[mask_masonry, "CostHead"] = "Masonry and plaster material only"
    t.loc[mask_concrete, "CostHead"] = "Concrete"
    return t

# ---------- Main CostHead Report Function ----------
def generate_costhead_report(gin_filepath: str, costhead_filepath: str, output_dir: str, match_mode: str = "contains") -> Dict[str, str]:
    """
    Generate CostHead reports based on the build_costhead_report_v9.py script

    Args:
        gin_filepath: Path to the GIN_Mapped Excel file
        costhead_filepath: Path to the CostHead Excel file
        output_dir: Directory to save the output reports
        match_mode: "contains" or "exact" for matching mode

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

    # Removed global SubProject filtering; filtering is enforced only for Steel/Concrete/Masonry within classifier

    mapping = load_mapping(costhead_raw)
    mapping = mapping[mapping["CostHead"] != ""].copy()

    # 1) Apply mapping (with special heads pre-assigned)
    tagged, map_expanded = assign_cost_head(gin, mapping, match_mode=match_mode)

    # 1b) Apply ParentWBS-based fallbacks so these are not counted as unmatched
    tagged = apply_parentwbs_fallbacks(tagged)

    # 2) Reallocate any remaining "Other" by SubProject rules
    tagged_final = reallocate_unmatched_by_subproject(tagged)

    # Canonicalize CostHead labels to avoid duplicates like STEEL vs Steel
    tagged_final["CostHead"] = tagged_final["CostHead"].apply(_canonicalize_costhead)


    # 3) Build reports from the final classification
    breakdown = (
        tagged_final.groupby(["CostHead","ActivityName","ParentWBS","SubProject","Project"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["CostHead","ActivityName","ParentWBS","SubProject","Project"])
    )

    matched_final = tagged_final.copy()  # everything, since we've allocated
    amenities_mask_summary = (
        (matched_final["CostHead"] == "Other") &
        ((matched_final["ActivityName"].fillna("") == "") | (matched_final["ParentWBS"].fillna("") == ""))
    )
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

    item_text: Dict[str, str] = {}
    allowed_item_heads = {"Steel", "Concrete", "Masonry and plaster material only"}
    for head in summary["CostHead"].tolist():
        if head not in allowed_item_heads:
            continue
        subset = matched_final[matched_final["CostHead"] == head]
        uniq = sorted(set(_token(r) for _, r in subset.iterrows()))
        MAX_ITEMS = 200
        if len(uniq) > MAX_ITEMS:
            shown = uniq[:MAX_ITEMS]
            more = len(uniq) - MAX_ITEMS
            item_text[head] = "; ".join(shown) + f"; ... (+{more} more)"
        else:
            item_text[head] = "; ".join(uniq)
    summary["ITEM Remark"] = summary["CostHead"].map(item_text).fillna("")

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
            agg = summary.groupby("CostHead", as_index=False).agg({
                "Material TotalAmount": "sum",
                "ITEM Remark": lambda s: "; ".join([x for x in s if str(x).strip()])
            })
        else:
            agg = summary.groupby("CostHead", as_index=False).agg({
                "ITEM Remark": lambda s: "; ".join([x for x in s if str(x).strip()])
            })
        summary = agg

    # Pad missing heads and enforce order
    order_df = pd.DataFrame({"CostHead": desired_order})
    summary = order_df.merge(summary, on="CostHead", how="left")
    if "Material TotalAmount" in summary.columns:
        summary["Material TotalAmount"] = summary["Material TotalAmount"].fillna("")
    summary["ITEM Remark"] = summary["ITEM Remark"].fillna("")

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

    # For audit, show what *would have been* unmatched before we reallocated
    other_before = tagged[tagged["CostHead"] == "Other"].copy()

    # Amenities: rows without ActivityName or ParentWBS
    def _fmt_amenity_name(sp: str) -> str:
        s = _norm_sub(sp)
        if s.startswith("TOWER ") and len(s) == len("TOWER X"):
            return s.replace(" ", " - ") + " ( No Activity Code )"
        if s == "PODIUM":
            return "PODIUM ( No Activity Code )"
        return f"{s} ( No Activity Code )" if s else "( No Activity Code )"

    amenities_mask = (other_before["ActivityName"].fillna("") == "") | (other_before["ParentWBS"].fillna("") == "")
    amenities_rows = other_before[amenities_mask].copy()
    amenities_rows["Amenity"] = amenities_rows["SubProject"].apply(_fmt_amenity_name)

    def _token(row):
        ig = str(row.get("ItemGroup", "") or "").strip()
        idc = str(row.get("ItemDesc", "") or "").strip()
        if ig and idc:
            return f"{ig} | {idc}"
        return idc or ig or "(blank)"

    amenities_items = (
        amenities_rows.groupby("Amenity")
        .apply(lambda g: "; ".join(sorted(set(_token(r) for _, r in g.iterrows()))))
        .reset_index(name="ITEM Remark")
    )
    amenities_summary = (
        amenities_rows.groupby("Amenity", as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"Material TotalAmount"})
        .merge(amenities_items, on="Amenity", how="left")
        .sort_values("Amenity")
    )

    # Include amenities also in unmatched views
    unmatched_detail = (
        other_before[["ActivityName","ParentWBS","SubProject","Project","ItemGroup","ItemDesc","IssueQty","Amount"]]
        .copy().rename(columns={"Amount":"IssueAmt"})
        .sort_values(["SubProject","ActivityName","ParentWBS","Project"])
    )
    unmatched_by_sp = (
        other_before.groupby(["SubProject","ActivityName","ParentWBS","Project"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["SubProject","TotalAmount"], ascending=[True, False])
    )

    # Append amenities summary at end of CostHead_Summary
    amenities_as_costhead = amenities_summary.rename(columns={"Amenity":"CostHead"})[
        ["CostHead","Material TotalAmount","ITEM Remark"]
    ]
    amenities_header = pd.DataFrame({
        "CostHead": ["Extra Remaining -Unmatched"],
        "Material TotalAmount": [""],
        "ITEM Remark": [""]
    })
    # Track the position (Excel row) of the amenities header for formatting
    summary_base_len = len(summary)
    summary = pd.concat([summary, amenities_header, amenities_as_costhead], ignore_index=True)

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

    return {
        "output_file": str(output_file),
        "output_dir": str(output_path),
        "sheets": ["CostHead_Breakdown", "CostHead_Summary", "Mapping_Expanded", "Unmatched_Detail", "Unmatched_By_SubProject"]
    }
