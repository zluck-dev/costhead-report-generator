#!/usr/bin/env python3
"""
build_costhead_report_v9.py

What it does:
- Reads: GIN_Mapped (ERP export) + CostHead (your 3-column mapping).
- Applies mapping (contains/normalized by default, switchable to --match exact).
- Then, for any rows still "Other":
    * If SubProject == TOWER A  -> CostHead = "without activity code TOWER A"
    * If SubProject == TOWER B  -> CostHead = "without activity code TOWER B"
    * Else if SubProject not blank -> CostHead = SubProject (each subproject becomes its own head)
    * Else (no SubProject): leave as "Other" (very rare)

Outputs to a NEW file:
  * CostHead_Breakdown           (detail by CostHead + fields)
  * CostHead_Summary             (Totals + Sources for each CostHead)
  * Mapping_Expanded             (how rules expanded)
  * Unmatched_Detail             (now shows the rows that *would* have been unmatched, but after reallocation it’s just for reference and audit; includes ItemGroup, ItemDesc, IssueQty, IssueAmt)
  * Unmatched_By_SubProject      (same — audit view)

Run:
  python build_costhead_report_v9.py --input "export_Skyteria_Mapped.xlsx" --output "Skyteria_With_CostHead.xlsx"
"""

import argparse, sys, re
from pathlib import Path
import warnings
import pandas as pd

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
}
MAP_COMPOSITE_SYNS = {
    "CostHead":   ["cost heads","cost head","costhead","head"],
    "Value":      ["parent wbs / activity name / sub project","parent wbs/activity name/sub project","value","criteria","keyword"],
    "FromWhere":  ["from where","match in","field","where"],
}

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
    parts = re.split(r"[;,]", v or "")
    return [_norm_text(p) for p in parts if p.strip()]

def _field_for(where_text: str) -> str:
    w = (where_text or "").lower()
    if "activity" in w or "task" in w: return "ActivityName"
    if "wbs" in w:                      return "ParentWBS"
    if "sub" in w:                      return "SubProject"
    if "project" in w:                  return "Project"
    return "ActivityName"

# ---------- Loaders ----------
def load_gin(df_raw: pd.DataFrame) -> pd.DataFrame:
    gin = _detect_header(df_raw, key_words=("activity","wbs","sub","project","amount","issue","qty","item"))
    cols = _pick_cols(gin, GIN_SYNS)
    need = ["ActivityName","ParentWBS","SubProject","Amount"]
    miss = [c for c in need if not cols[c]]
    if not cols.get("Project"):
        gin["__ProjectBlank__"] = ""
        cols["Project"] = "__ProjectBlank__"
    if miss: raise SystemExit(f"GIN_Mapped missing columns: {miss}\nFound: {list(gin.columns)}")

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

    return out

def load_mapping(df_raw: pd.DataFrame) -> pd.DataFrame:
    mp = _detect_header(df_raw, key_words=("cost","head","activity","wbs","sub","from","project"))
    comp = _pick_cols(mp, MAP_COMPOSITE_SYNS)
    if not comp.get("CostHead") or not comp.get("Value") or not comp.get("FromWhere"):
        raise SystemExit("CostHead sheet not recognized. Expected: COST HEADS | parent WBS / Activity Name / Sub Project | From where")
    return pd.DataFrame({
        "CostHead":  mp[comp["CostHead"]].fillna("").astype(str).map(_norm_text),
        "ValueRaw":  mp[comp["Value"]].fillna("").astype(str),
        "FromWhere": mp[comp["FromWhere"]].fillna("").astype(str),
        "Value":     mp[comp["Value"]].fillna("").astype(str).map(_norm_text),
    })

# ---------- Mapping ----------
def assign_cost_head(gin: pd.DataFrame, mapping: pd.DataFrame, match_mode="contains"):
    tagged = gin.copy()
    tagged["CostHead"] = "Other"
    assigned_mask = pd.Series(False, index=tagged.index)

    expanded = []
    for _, r in mapping.iterrows():
        head = r["CostHead"]; field = _field_for(r["FromWhere"])
        for kw in _split_keywords(r["Value"]):
            if head and kw: expanded.append({"CostHead": head, "Field": field, "Keyword": kw})
    map_expanded = pd.DataFrame(expanded)

    if not map_expanded.empty:
        for _, mr in map_expanded.iterrows():
            field = mr["Field"]; kw = mr["Keyword"]; hay = tagged[field]
            mask = (hay == kw) if match_mode == "exact" else hay.str.contains(re.escape(kw), na=False)
            to_assign = mask & (~assigned_mask)
            if to_assign.any():
                tagged.loc[to_assign, "CostHead"] = mr["CostHead"]
                assigned_mask = assigned_mask | to_assign

    return tagged, map_expanded

# ---------- Reallocate unmatched by SubProject ----------
def reallocate_unmatched_by_subproject(tagged: pd.DataFrame) -> pd.DataFrame:
    t = tagged.copy()
    is_other = (t["CostHead"] == "Other")

    # normalize subproject for comparison
    sp = t["SubProject"].fillna("").astype(str)
    sp_norm = sp.str.upper().str.replace(r"\s+", " ", regex=True)

    mask_tower_a = is_other & sp_norm.isin(["TOWER A", "TOWER - A"])
    mask_tower_b = is_other & sp_norm.isin(["TOWER B", "TOWER - B"])

    t.loc[mask_tower_a, "CostHead"] = "WITHOUT ACTIVITY CODE TOWER A"
    t.loc[mask_tower_b, "CostHead"] = "WITHOUT ACTIVITY CODE TOWER B"

    # remaining "Other" with a non-empty SubProject -> use SubProject name as CostHead
    mask_sp_rest = is_other & (sp_norm != "")
    t.loc[mask_sp_rest, "CostHead"] = sp[mask_sp_rest]  # keep original (human-readable) SubProject as the head

    return t

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="CostHead report with SubProject-based allocation of 'Other'")
    ap.add_argument("--input", required=True, help="Excel with GIN_Mapped and CostHead")
    ap.add_argument("--output", required=True, help="NEW Excel path to write")
    ap.add_argument("--gin-sheet", default="GIN_Mapped")
    ap.add_argument("--map-sheet", default="CostHead")
    ap.add_argument("--match", choices=["contains","exact"], default="contains")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists(): sys.exit(f"Input not found: {src}")

    xl = pd.ExcelFile(src, engine="openpyxl")
    if args.gin_sheet not in xl.sheet_names: sys.exit(f"GIN sheet '{args.gin_sheet}' not found. Sheets: {xl.sheet_names}")
    if args.map_sheet not in xl.sheet_names: sys.exit(f"Mapping sheet '{args.map_sheet}' not found. Sheets: {xl.sheet_names}")

    gin_raw = xl.parse(args.gin_sheet, header=None)
    map_raw = xl.parse(args.map_sheet, header=None)

    gin = load_gin(gin_raw)
    mapping = load_mapping(map_raw)
    mapping = mapping[mapping["CostHead"] != ""].copy()

    # 1) Apply mapping
    tagged, map_expanded = assign_cost_head(gin, mapping, match_mode=args.match)

    # 2) Reallocate any remaining "Other" by SubProject rules you requested
    tagged_final = reallocate_unmatched_by_subproject(tagged)

    # 3) Build reports from the final classification
    breakdown = (
        tagged_final.groupby(["CostHead","ActivityName","ParentWBS","SubProject","Project"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values(["CostHead","ActivityName","ParentWBS","SubProject","Project"])
    )

    matched_final = tagged_final.copy()  # everything, since we've allocated
    summary = (
        matched_final.groupby("CostHead", as_index=False)["Amount"]
        .sum().rename(columns={"Amount":"TotalAmount"})
        .sort_values("CostHead")
    )

    def make_sources(g: pd.DataFrame) -> str:
        tuples = (
            g[["ActivityName","ParentWBS","SubProject","Project"]]
            .drop_duplicates().astype(str)
            .agg(" | ".join, axis=1).tolist()
        )
        return "; ".join(tuples)[:60000]

    srcs = (
        matched_final.groupby("CostHead")[["ActivityName","ParentWBS","SubProject","Project"]]
        .apply(make_sources)
        .reset_index(name="Sources")
    )
    summary = summary.merge(srcs, on="CostHead", how="left")

    # For audit, show what *would have been* unmatched before we reallocated
    other_before = tagged[tagged["CostHead"] == "Other"].copy()
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

    # 4) Write NEW workbook
    out = Path(args.output)
    with pd.ExcelWriter(out, engine="openpyxl", mode="w") as xw:
        breakdown.to_excel(xw, sheet_name="CostHead_Breakdown", index=False)
        summary.to_excel(xw, sheet_name="CostHead_Summary", index=False)
        map_expanded.to_excel(xw, sheet_name="Mapping_Expanded", index=False)
        unmatched_detail.to_excel(xw, sheet_name="Unmatched_Detail", index=False)
        unmatched_by_sp.to_excel(xw, sheet_name="Unmatched_By_SubProject", index=False)

    print(f"Done. Wrote reports to: {out.resolve()}")
    print("Sheets: CostHead_Breakdown, CostHead_Summary, Mapping_Expanded, Unmatched_Detail, Unmatched_By_SubProject")

if __name__ == "__main__":
    main()
