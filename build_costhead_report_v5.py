#!/usr/bin/env python3
"""
build_costhead_report_v5.py
- Reads: GIN_Mapped (data) + CostHead (3-column mapping you maintain)
- Matches by the field in "From where" (Activity / Parent WBS / Sub Project / Project)
- Matching is case-insensitive, space/–/—/slash-normalized, "contains" by default
- Outputs to a NEW FILE:
  * CostHead_Breakdown
  * CostHead_Summary  (matched cost heads) + APPENDS:
      - Unmatched (by SubProject)
      - Unmatched (by ActivityName + ParentWBS) for rows without SubProject

Run example:
  python build_costhead_report_v5.py --input "export_Skyteria_Mapped.xlsx" --output "Skyteria_With_CostHead.xlsx"
Options:
  --gin-sheet GIN_Mapped   --map-sheet CostHead   --match contains|exact
"""

import argparse, sys, re
from pathlib import Path
import pandas as pd

# ---------- Synonyms (auto-detect headers) ----------
GIN_SYNS = {
    "ActivityName": ["activity name","activity_name","task name","task","act name"],
    "ParentWBS":    ["parent wbs","wbs name","parent wbs name","wbs parent","wbs level 1","wbs"],
    "SubProject":   ["sub project","subproject","sub_project","location","subproject name","sub proj","sub_proj"],
    "Project":      ["project","project name","project_name"],
    "Amount":       ["gin issue amt","issue amount","amount","total issued amount","issued amount","value","net amount","total amount"],
}
MAP_COMPOSITE_SYNS = {
    "CostHead":   ["cost heads","cost head","costhead","head"],
    "Value":      ["parent wbs / activity name / sub project","parent wbs/activity name/sub project","value","criteria","keyword"],
    "FromWhere":  ["from where","match in","field","where"],
}

# ---------- Normalization ----------
def _clean_name(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().lower()
    s = re.sub(r'[_\-]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s

def _norm_text(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = s.replace("\u00a0"," ")           # non-breaking space -> space
    s = s.replace("–","-").replace("—","-")
    s = s.replace("/", " ")               # normalize separators
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().upper()

def _pick_col(df: pd.DataFrame, options):
    clean = {_clean_name(c): c for c in df.columns}
    for cand in options:
        k = _clean_name(cand)
        if k in clean:
            return clean[k]
        for kk, orig in clean.items():
            if k in kk:
                return orig
    return ""

def _pick_cols(df: pd.DataFrame, spec: dict):
    out = {}
    for want, alts in spec.items():
        out[want] = _pick_col(df, [want] + alts)
    return out

def _detect_header(df_raw: pd.DataFrame, key_words=("cost","head","activity","wbs","sub","project","amount"), scan=20):
    df = df_raw.copy()
    if len(df) == 0:
        return df
    upto = min(scan, len(df))
    for r in range(upto):
        row = [str(x).strip() for x in df.iloc[r].tolist()]
        row_text = " ".join(x.lower() for x in row)
        if any(k in row_text for k in key_words):
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
    return [ _norm_text(p) for p in parts if p.strip() ]

def load_gin(df_raw: pd.DataFrame) -> pd.DataFrame:
    gin = _detect_header(df_raw, key_words=("activity","wbs","sub","project","amount","issue"))
    cols = _pick_cols(gin, GIN_SYNS)
    need = ["ActivityName","ParentWBS","SubProject","Amount"]
    miss = [c for c in need if not cols[c]]
    # Project optional
    if not cols.get("Project"):
        gin["__ProjectBlank__"] = ""
        cols["Project"] = "__ProjectBlank__"
    if miss:
        raise SystemExit(f"GIN_Mapped missing columns: {miss}\nFound: {list(gin.columns)}")
    out = pd.DataFrame({
        "ActivityName": _norm_series_text(gin[cols["ActivityName"]]),
        "ParentWBS":    _norm_series_text(gin[cols["ParentWBS"]]),
        "SubProject":   _norm_series_text(gin[cols["SubProject"]]),
        "Project":      _norm_series_text(gin[cols["Project"]]),
        "Amount":       _norm_series_amt(gin[cols["Amount"]]),
    })
    return out

def load_mapping(df_raw: pd.DataFrame) -> pd.DataFrame:
    mp = _detect_header(df_raw, key_words=("cost","head","activity","wbs","sub","from","project"))
    comp = _pick_cols(mp, MAP_COMPOSITE_SYNS)
    if not comp.get("CostHead") or not comp.get("Value") or not comp.get("FromWhere"):
        raise SystemExit(
            "CostHead sheet not recognized. Expected:\n"
            "  COST HEADS | parent WBS / Activity Name / Sub Project | From where"
        )
    return pd.DataFrame({
        "CostHead":  mp[comp["CostHead"]].fillna("").astype(str).map(_norm_text),
        "ValueRaw":  mp[comp["Value"]].fillna("").astype(str),
        "FromWhere": mp[comp["FromWhere"]].fillna("").astype(str),
        "Value":     mp[comp["Value"]].fillna("").astype(str).map(_norm_text),
    })

def _field_for(where_text: str) -> str:
    w = (where_text or "").lower()
    if "activity" in w or "task" in w:      return "ActivityName"
    if "wbs" in w:                           return "ParentWBS"
    if "sub" in w:                           return "SubProject"
    if "project" in w:                       return "Project"
    return "ActivityName"

def assign_cost_head(gin: pd.DataFrame, mapping: pd.DataFrame, match_mode="contains"):
    tagged = gin.copy()
    tagged["CostHead"] = "Other"
    assigned_mask = pd.Series(False, index=tagged.index)

    # expand mapping to one row per keyword
    expanded = []
    for _, r in mapping.iterrows():
        head = r["CostHead"]
        field = _field_for(r["FromWhere"])
        for kw in _split_keywords(r["Value"]):
            if head and kw:
                expanded.append({"CostHead": head, "Field": field, "Keyword": kw})
    map_expanded = pd.DataFrame(expanded)

    # apply first-match-wins in sheet order
    if not map_expanded.empty:
        for _, mr in map_expanded.iterrows():
            field = mr["Field"]; kw = mr["Keyword"]
            hay = tagged[field]
            mask = (hay == kw) if match_mode == "exact" else hay.str.contains(re.escape(kw), na=False)
            to_assign = mask & (~assigned_mask)
            if to_assign.any():
                tagged.loc[to_assign, "CostHead"] = mr["CostHead"]
                assigned_mask = assigned_mask | to_assign

    return tagged, map_expanded

def main():
    ap = argparse.ArgumentParser(description="CostHead reports with unmatched appended to summary")
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

    tagged, map_expanded = assign_cost_head(gin, mapping, match_mode=args.match)

    # Normal breakdown (includes "Other" rows too but we'll mostly use for matched)
    breakdown = (tagged.groupby(["CostHead","ActivityName","ParentWBS","SubProject","Project"], as_index=False)["Amount"]
                 .sum().rename(columns={"Amount":"TotalAmount"})
                 .sort_values(["CostHead","ActivityName","ParentWBS","SubProject","Project"]))

    # Matched summary (exclude "Other")
    matched = tagged[tagged["CostHead"] != "Other"]
    summary = (matched.groupby("CostHead", as_index=False)["Amount"]
               .sum().rename(columns={"Amount":"TotalAmount"})
               .sort_values("CostHead"))

    # ---- Append unmatched totals ----
    other = tagged[tagged["CostHead"] == "Other"].copy()

    # 1) By SubProject (non-empty)
    um_by_sp = (other[other["SubProject"] != ""]
                .groupby("SubProject", as_index=False)["Amount"].sum()
                .rename(columns={"SubProject":"Label","Amount":"TotalAmount"}))
    um_by_sp["Label"] = "Unmatched (SubProject: " + um_by_sp["Label"] + ")"

    # 2) By Activity+ParentWBS where SubProject is blank
    tmp = other[other["SubProject"] == ""].copy()
    tmp["Pair"] = tmp["ActivityName"].fillna("") + " | " + tmp["ParentWBS"].fillna("")
    um_by_ap = (tmp.groupby("Pair", as_index=False)["Amount"].sum()
                  .rename(columns={"Pair":"Label","Amount":"TotalAmount"}))
    um_by_ap["Label"] = "Unmatched (Activity+WBS: " + um_by_ap["Label"] + ")"

    # Combine unmatched blocks
    unmatched_rows = pd.concat([um_by_sp, um_by_ap], ignore_index=True)
    if not unmatched_rows.empty:
        unmatched_rows = unmatched_rows.sort_values("Label")
        # Append to summary (as extra rows at the bottom)
        extra = unmatched_rows.rename(columns={"Label":"CostHead"})
        summary = pd.concat([summary, pd.DataFrame([{"CostHead":"", "TotalAmount":""}]), extra], ignore_index=True)

    out = Path(args.output)
    with pd.ExcelWriter(out, engine="openpyxl", mode="w") as xw:
        breakdown.to_excel(xw, sheet_name="CostHead_Breakdown", index=False)
        summary.to_excel(xw, sheet_name="CostHead_Summary", index=False)
        map_expanded.to_excel(xw, sheet_name="Mapping_Expanded", index=False)

    print(f"Done. Wrote reports to: {out.resolve()}")
    print("Sheets: CostHead_Breakdown, CostHead_Summary, Mapping_Expanded")

if __name__ == "__main__":
    main()
