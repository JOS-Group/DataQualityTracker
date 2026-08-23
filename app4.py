"""
DART - Data Assurance Reconciliation Tracker
One-file FastAPI website edition

Run:
  python -m pip install fastapi uvicorn pandas numpy plotly openpyxl python-multipart openai
  python -m uvicorn app3:app --reload

What is included in this update:
- Metric Explanations page and inline metric help cards
- Editable Lineage & Mapping Catalog with add, edit, delete, and seed-from-data support
- Impact Explorer with field-to-report/KPI/owner mapping and impact scoring
- Sankey Lineage Visualization for source table -> source field -> target field -> report/KPI
- Seeded Workflow Issues so Action Center is populated immediately
- Governance Center with controls, decisions, owners, rule templates, and export-ready summary
- Better empty states so charts do not appear blank when filters return no rows
- Revised page names, plain-English copy, and nonduplicative page content
- Real-workbook-only calculations and stronger chart validation guards
- Clear filter selection labels and safer persona fallbacks
- Loads the real All Streams reconciliation workbook by default
- Enriches each reconciliation row from the CMS NCH-to-STTM mapping workbook
- One Python file only: FastAPI backend + embedded HTML/CSS/JS frontend

Prototype note:
- This is a local demo/prototype app. Do not use with sensitive production data without
  replacing prototype state/auth patterns with production-grade storage and authentication.
"""
from __future__ import annotations

import io
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

try:
    from openai import OpenAI
except Exception:  # pragma: no cover
    OpenAI = None

APP_VERSION = "7.1.0"
app = FastAPI(title="DART - Data Assurance Reconciliation Tracker", version=APP_VERSION)

USERS_FILE = Path(__file__).resolve().parent / "dart_users.json"

STATE: Dict[str, Any] = {
    "persona": None,
    "df": pd.DataFrame(),
    "source": "",
    "issues": [],
    "connections": [],
    "impacts": [],
    "governance": [],
    "rules": [],
    "chat": [],
    "snapshots": [],
    "current_user": None,
}

# -----------------------------------------------------------------------------
# User/profile helpers
# -----------------------------------------------------------------------------

def load_users() -> Dict[str, Any]:
    if not USERS_FILE.exists():
        return {"users": {}}
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"users": {}}


def save_users(data: Dict[str, Any]) -> None:
    USERS_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")


def public_user(username: Optional[str]) -> Optional[Dict[str, Any]]:
    if not username:
        return None
    user = load_users().get("users", {}).get(username)
    if not user:
        return None
    profile = user.get("profile", {})
    return {
        "username": username,
        "display_name": profile.get("display_name", username),
        "email": profile.get("email", ""),
        "organization": profile.get("organization", ""),
        "role": profile.get("role", ""),
        "created_at": user.get("created_at", "local prototype"),
        "storage_file": USERS_FILE.name,
    }

# -----------------------------------------------------------------------------
# Data helpers
# -----------------------------------------------------------------------------

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def find_col(df: pd.DataFrame, choices: List[str]) -> Optional[str]:
    exact = {str(c).lower().strip(): c for c in df.columns}
    for choice in choices:
        key = choice.lower().strip()
        if key in exact:
            return exact[key]
    for col in df.columns:
        for choice in choices:
            if choice.lower() in str(col).lower():
                return col
    return None


def standardize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = norm_cols(df)
    aliases = {
        "Stream": ["Stream", "Claim Stream", "Domain", "LOB", "Line of Business"],
        "NCH Target Table": ["NCH Target Table", "NCH Table", "System A Table", "Source Table", "Source Entity"],
        "NCH Target Column": ["NCH Target Column", "NCH Column", "System A Column", "Source Column", "Field"],
        "SS Table": ["SS Table", "Shared Systems Table", "System B Table", "Target Table", "Target Entity"],
        "SS Column": ["SS Column", "SS Colmn if different", "Shared Systems Column", "System B Column", "Target Column"],
        "MatchedClaims": ["MatchedClaims", "Matched Claims", "1/1/24 - 12/31/24 Matched claims", "Matched Events", "Matched", "Match Count"],
        "NotMatchedClaims": ["NotMatchedClaims", "Not Matched Claims", "1/1/24 - 12/31/24 Not Matched claims", "Not-Matched Claims", "Unmatched", "Mismatch Count"],
        "MatchRate": ["MatchRate", "Match Rate", "% matched", "Match %", "Match Percentage", "Accuracy"],
        "Classification": ["Classification", "ClassificationClean", "Class", "Status"],
        "Sub-Classification": ["Sub-Classification", "Sub Classification", "Subclass", "Reason", "Driver"],
        "Disposition": ["Disposition", "DispositionClean", "Action"],
        "Recommendation": ["Recommendation", "Recommended Action", "Next Step"],
        "Comments": ["Comments", "Comments/Examples", "Comment", "Notes"],
        "Lynette Recommendation": ["Lynette Recommendation", "Lynette's recommendation", "Lynnette's recommendation"],
        "Owner": ["Owner", "Business Owner", "Team"],
        "Report": ["Report", "Dashboard", "Downstream Report"],
        "KPI": ["KPI", "Metric", "Measure"],
    }
    rename_map = {}
    for target, options in aliases.items():
        col = find_col(df, options)
        if col and col != target:
            rename_map[col] = target
    df = df.rename(columns=rename_map)

    for col in ["MatchedClaims", "NotMatchedClaims", "MatchRate"]:
        if col in df:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
    if "MatchedClaims" not in df:
        df["MatchedClaims"] = np.nan
    if "NotMatchedClaims" not in df:
        df["NotMatchedClaims"] = np.nan
    if "MatchRate" in df and df["MatchRate"].dropna().size and df["MatchRate"].dropna().max() > 1.5:
        df["MatchRate"] = df["MatchRate"] / 100
    if "MatchRate" not in df:
        total = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
        df["MatchRate"] = np.where(total > 0, df["MatchedClaims"] / total, np.nan)
    if "Classification" not in df:
        df["Classification"] = np.where(df["MatchRate"] >= 0.99, "Match", np.where(df["MatchRate"].notna(), "Different", "Missing"))

    for col in [
        "Stream", "NCH Target Table", "NCH Target Column", "SS Table", "SS Column",
        "Disposition", "Recommendation", "Comments", "Lynette Recommendation", "Sub-Classification", "Owner", "Report", "KPI",
        "Source Field Name", "Source Field Description", "Business Rule", "Mapping Comments", "Mapping Notes", "Mapping Status",
    ]:
        if col not in df:
            df[col] = ""

    df["Comparable"] = df["Classification"].astype(str).str.lower().str.strip().isin(["match", "different", "mismatch"])
    disp = df["Disposition"].astype(str).str.lower()
    cls = df["Classification"].astype(str).str.lower()
    df["NeedsChange"] = disp.str.contains("action|change|review|fix|remediate", na=False) | cls.isin(["missing", "different", "mismatch"])
    df["TotalClaims"] = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
    df["ImpactScore"] = ((1 - df["MatchRate"].fillna(0)) * np.log1p(df["TotalClaims"].fillna(0))).round(2)
    df["RiskTier"] = np.select(
        [df["ImpactScore"] >= 6, df["ImpactScore"] >= 3, df["ImpactScore"] > 0],
        ["Critical", "Elevated", "Watch"],
        default="Stable",
    )
    return df


def pct(x: Any) -> str:
    return "-" if pd.isna(x) else f"{float(x):.1%}"


def fmt_int(x: Any) -> str:
    try:
        return f"{int(float(x)):,}"
    except Exception:
        return "0"


def field_average(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    x = df[df["Comparable"] & df["MatchRate"].notna()]
    return float(x["MatchRate"].mean()) if len(x) else np.nan


def weighted_rate(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    m = float(df["MatchedClaims"].fillna(0).sum())
    u = float(df["NotMatchedClaims"].fillna(0).sum())
    return m / (m + u) if (m + u) else np.nan


def _clean_key(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip()).upper()
    return "" if text in {"", "NAN", "NONE", "-"} else text


def _join_unique(values: pd.Series) -> str:
    seen: List[str] = []
    for value in values:
        text = str(value).strip() if pd.notna(value) else ""
        if text and text.lower() not in {"nan", "none", "-"} and text not in seen:
            seen.append(text)
    return " | ".join(seen)


def load_mapping_catalog(path: Path) -> pd.DataFrame:
    mapping = pd.read_excel(path, sheet_name="Lynnette's Version", header=1, engine="openpyxl")
    mapping = norm_cols(mapping)
    needed = ["NCH Target Table", "NCH Target Column"]
    if not all(c in mapping.columns for c in needed):
        raise ValueError("Mapping workbook is missing NCH Target Table/NCH Target Column.")
    mapping = mapping[mapping["NCH Target Table"].notna() & mapping["NCH Target Column"].notna()].copy()
    mapping["_table_key"] = mapping["NCH Target Table"].map(_clean_key)
    mapping["_column_key"] = mapping["NCH Target Column"].map(_clean_key)
    mapping = mapping[(mapping["_table_key"] != "") & (mapping["_column_key"] != "")]
    rename = {
        "NCH Cobol Copy book Field Name": "Source Field Description",
        "In Table Source Fields": "Source Field Name",
        "Business Rule": "Business Rule",
        "Comments": "Mapping Comments",
        "Notes": "Mapping Notes",
        "Data Match? Y/N": "Mapping Status",
    }
    mapping = mapping.rename(columns={k: v for k, v in rename.items() if k in mapping.columns})
    keep = ["_table_key", "_column_key"] + [c for c in rename.values() if c in mapping.columns]
    mapping = mapping[keep]
    agg = {c: _join_unique for c in keep if c not in {"_table_key", "_column_key"}}
    return mapping.groupby(["_table_key", "_column_key"], as_index=False).agg(agg)


def enrich_with_mapping(df: pd.DataFrame, mapping_path: Path) -> pd.DataFrame:
    if df.empty or not mapping_path.exists():
        return df
    left = df.copy()
    left["_table_key"] = left["NCH Target Table"].map(_clean_key)
    left["_column_key"] = left["NCH Target Column"].map(_clean_key)
    try:
        catalog = load_mapping_catalog(mapping_path)
        left = left.merge(catalog, on=["_table_key", "_column_key"], how="left")
    except Exception as exc:
        # Mapping enrichment is optional. Preserve the reconciliation data if the
        # mapping workbook has a different sheet/header layout.
        print(f"WARNING: Mapping enrichment skipped: {exc}")
        left["Mapping Load Warning"] = str(exc)

    source = left.get("Source Field Name", pd.Series("", index=left.index)).fillna("").astype(str).str.strip()
    rule = left.get("Business Rule", pd.Series("", index=left.index)).fillna("").astype(str).str.strip()
    left["Mapping Found"] = source.ne("") | rule.ne("")
    return left.drop(columns=["_table_key", "_column_key"], errors="ignore")


def load_local_data() -> Tuple[pd.DataFrame, str]:
    here = Path(__file__).resolve().parent
    data_path = here / "All Streams June 2026.xlsx"
    mapping_path = here / "CMS Version NCH to NCH STTM 12_13_2024.xlsx"
    if not data_path.exists():
        return pd.DataFrame(), f"Missing required file: {data_path.name}"
    try:
        df = pd.read_excel(data_path, sheet_name="Sheet1", engine="openpyxl")
        df = standardize(df)
        df = enrich_with_mapping(df, mapping_path)
        df = standardize(df)
        source = data_path.name
        if mapping_path.exists():
            source += f" + {mapping_path.name}"
        else:
            source += " (mapping workbook not found)"
        return df, source
    except Exception as exc:
        return pd.DataFrame(), f"Could not load workbook data: {exc}"


def seed_lineage_from_df(df: pd.DataFrame) -> None:
    if STATE["connections"]:
        return
    sample = df.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(12)
    records = []
    for i, r in sample.iterrows():
        records.append({
            "id": f"L{i}",
            "Stream": str(r.get("Stream", "")),
            "SourceTable": str(r.get("NCH Target Table", "")),
            "SourceField": str(r.get("NCH Target Column", "")),
            "TargetTable": str(r.get("SS Table", "")),
            "TargetField": str(r.get("SS Column", "")),
            "Report": str(r.get("Report", "")) or f"{r.get('Stream', 'Stream')} Quality Report",
            "KPI": str(r.get("KPI", "")) or f"{r.get('Stream', 'Stream')} Match KPI",
            "Owner": str(r.get("Owner", "")) or "Data Quality",
            "Status": "Draft" if r.get("RiskTier") in ["Critical", "Elevated"] else "Validated",
            "RiskTier": str(r.get("RiskTier", "")),
            "Notes": str(r.get("Recommendation", "")),
        })
    STATE["connections"] = records


def seed_workflow_items(df: pd.DataFrame) -> None:
    if STATE["issues"]:
        return
    sample = df[df["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(8)
    items = []
    lane_cycle = ["Open", "In progress", "Open", "Blocked", "Open", "In progress", "Resolved", "Open"]
    for pos, (_, r) in enumerate(sample.iterrows()):
        items.append({
            "id": f"I{pos+1}",
            "Stream": str(r["Stream"]),
            "Field": str(r["NCH Target Column"]),
            "RiskTier": str(r["RiskTier"]),
            "Priority": "Critical" if r["RiskTier"] == "Critical" else "High" if r["RiskTier"] == "Elevated" else "Medium",
            "Owner": str(r.get("Owner", "Data Quality")) or "Data Quality",
            "Status": lane_cycle[pos % len(lane_cycle)],
            "Due": "Next review cycle",
            "Note": str(r.get("Recommendation", "Confirm mapping and lineage.")),
        })
    STATE["issues"] = items


def seed_impacts(df: pd.DataFrame) -> None:
    if STATE["impacts"]:
        return
    sample = df.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(8)
    STATE["impacts"] = [{
        "id": f"M{i+1}",
        "Stream": str(r["Stream"]),
        "Field": str(r["NCH Target Column"]),
        "Report": str(r.get("Report", "")) or f"{r['Stream']} Quality Report",
        "KPI": str(r.get("KPI", "")) or f"{r['Stream']} Match KPI",
        "Owner": str(r.get("Owner", "")) or "Reporting",
        "Impact": "Critical" if r["RiskTier"] == "Critical" else "High" if r["RiskTier"] == "Elevated" else "Moderate",
        "DecisionNeed": "Confirm remediation priority" if r["RiskTier"] in ["Critical", "Elevated"] else "Monitor",
    } for i, (_, r) in enumerate(sample.iterrows())]


def seed_governance() -> None:
    if STATE["governance"]:
        return
    STATE["governance"] = [
        {"id": "G1", "Control": "Critical mapping owner assigned", "Owner": "Data Quality", "Cadence": "Weekly", "Status": "Active", "Evidence": "Action Center owner and note"},
        {"id": "G2", "Control": "Lineage status reviewed", "Owner": "Reporting", "Cadence": "Monthly", "Status": "Draft", "Evidence": "Lineage catalog status"},
        {"id": "G3", "Control": "Executive KPI impact confirmed", "Owner": "Leadership", "Cadence": "Review cycle", "Status": "Active", "Evidence": "Impact Explorer mapping"},
        {"id": "G4", "Control": "Automation threshold defined", "Owner": "Engineering", "Cadence": "As needed", "Status": "Planned", "Evidence": "Automation rule template"},
    ]


def ensure_data() -> None:
    if STATE["df"].empty:
        df, name = load_local_data()
        if df.empty:
            raise RuntimeError(name or "The required All Streams workbook could not be loaded.")
        STATE["df"] = df
        STATE["source"] = name
    df = standardize(STATE["df"])


def custom_mask(series: pd.Series, op: str, value: str) -> pd.Series:
    s = series.astype(str)
    value_text = str(value)
    op = str(op or "contains")
    if op == "equals":
        return s.str.lower() == value_text.lower()
    if op == "not_equals":
        return s.str.lower() != value_text.lower()
    if op == "starts_with":
        return s.str.lower().str.startswith(value_text.lower(), na=False)
    if op == "ends_with":
        return s.str.lower().str.endswith(value_text.lower(), na=False)
    if op in ["gt", "gte", "lt", "lte"]:
        nums = pd.to_numeric(series, errors="coerce")
        try:
            v = float(value_text)
        except Exception:
            return pd.Series([True] * len(series), index=series.index)
        if op == "gt":
            return nums > v
        if op == "gte":
            return nums >= v
        if op == "lt":
            return nums < v
        if op == "lte":
            return nums <= v
    return s.str.contains(value_text, case=False, na=False)


def apply_filters(filters: Dict[str, Any]) -> pd.DataFrame:
    ensure_data()
    df = standardize(STATE["df"])
    fdf = df.copy()
    streams = filters.get("streams")
    classes = filters.get("classes")
    tiers = filters.get("tiers")
    reasons = filters.get("reasons")
    min_rate = float(filters.get("min_rate", 0) or 0)
    min_impact = float(filters.get("min_impact", 0) or 0)
    min_unmatched = float(filters.get("min_unmatched", 0) or 0)
    actions_only = bool(filters.get("actions_only", False))
    search = str(filters.get("search", "")).strip()
    custom_filters = filters.get("custom_filters") or []
    if streams:
        fdf = fdf[fdf["Stream"].astype(str).isin(streams)]
    if classes:
        fdf = fdf[fdf["Classification"].astype(str).isin(classes)]
    if tiers:
        fdf = fdf[fdf["RiskTier"].astype(str).isin(tiers)]
    if reasons:
        fdf = fdf[fdf["Sub-Classification"].astype(str).isin(reasons)]
    if min_rate > 0:
        fdf = fdf[fdf["MatchRate"].isna() | (fdf["MatchRate"] >= min_rate)]
    if min_impact > 0:
        fdf = fdf[fdf["ImpactScore"].fillna(0) >= min_impact]
    if min_unmatched > 0:
        fdf = fdf[fdf["NotMatchedClaims"].fillna(0) >= min_unmatched]
    if actions_only:
        fdf = fdf[fdf["NeedsChange"]]
    if search:
        mask = fdf.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        fdf = fdf[mask]
    for cf in custom_filters:
        field = str(cf.get("field", "")).strip()
        op = str(cf.get("op", "contains")).strip()
        value = str(cf.get("value", "")).strip()
        if field in fdf.columns and value:
            fdf = fdf[custom_mask(fdf[field], op, value)]
    return fdf


def clean_records(df: pd.DataFrame, limit: int = 200) -> List[Dict[str, Any]]:
    return df.head(limit).replace({np.nan: None}).to_dict("records")


def chart_html(fig: go.Figure, height: int = 360) -> str:
    fig.update_layout(
        template="plotly_dark",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"color": "#fff7d1", "family": "Inter, Segoe UI, Arial"},
        colorway=["#f9d976", "#d4af37", "#a87918", "#f2c14e", "#8a6a1f", "#ffe8a3"],
        margin={"l": 44, "r": 18, "t": 58, "b": 42},
        legend={"orientation": "h", "y": -0.2},
    )
    return pio.to_html(fig, include_plotlyjs=False, full_html=False, config={"displayModeBar": False, "responsive": True})


def no_data_chart(title: str, message: str = "No data is available for the current selection.") -> str:
    fig = go.Figure()
    fig.add_annotation(
        text=message,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        align="center",
        font={"size": 15, "color": "#f9e6a2"},
    )
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    fig.update_layout(title=title)
    return chart_html(fig)


def summary_payload(fdf: pd.DataFrame) -> Dict[str, Any]:
    if fdf.empty:
        return {"rows": 0, "field_average": "-", "weighted_rate": "-", "actions": "0", "streams": "0", "unmatched": "0", "risk_count": "0", "health_score": "-", "critical": "0", "risk_score": "-", "affected_reports": "0", "top_actions": [], "stream_summary": []}
    actions = fdf[fdf["NeedsChange"]]
    health = weighted_rate(fdf)
    risk_score = float(fdf["ImpactScore"].fillna(0).mean()) if len(fdf) else np.nan
    stream_summary = fdf.groupby("Stream", as_index=False).agg(
        Fields=("MatchRate", "count"),
        AverageMatch=("MatchRate", "mean"),
        Matched=("MatchedClaims", "sum"),
        NotMatched=("NotMatchedClaims", "sum"),
        AverageImpact=("ImpactScore", "mean"),
        Critical=("RiskTier", lambda x: int((x == "Critical").sum())),
    )
    stream_summary["WeightedMatch"] = stream_summary["AverageMatch"]
    stream_summary_records = stream_summary.sort_values("NotMatched", ascending=False).replace({np.nan: None}).to_dict("records")
    top_actions = actions.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(15).replace({np.nan: None}).to_dict("records")
    affected = max(len(STATE.get("impacts", [])), fdf["Report"].replace("", np.nan).nunique(dropna=True), fdf["Stream"].nunique())
    return {
        "rows": len(fdf),
        "field_average": pct(field_average(fdf)),
        "weighted_rate": pct(weighted_rate(fdf)),
        "actions": fmt_int(actions.shape[0]),
        "streams": fmt_int(fdf["Stream"].nunique()),
        "unmatched": fmt_int(fdf["NotMatchedClaims"].fillna(0).sum()),
        "risk_count": fmt_int((fdf["RiskTier"].isin(["Critical", "Elevated"])).sum()),
        "critical": fmt_int((fdf["RiskTier"] == "Critical").sum()),
        "risk_score": "-" if pd.isna(risk_score) else f"{risk_score:.1f}",
        "affected_reports": fmt_int(affected),
        "health_score": "-" if pd.isna(health) else f"{health * 100:.0f}",
        "top_actions": top_actions,
        "stream_summary": stream_summary_records,
    }


def add_snapshot(fdf: pd.DataFrame) -> None:
    if fdf.empty:
        return
    snap = {"Label": f"S{len(STATE['snapshots']) + 1}", "FieldAverage": field_average(fdf), "WeightedRate": weighted_rate(fdf), "Critical": int((fdf["RiskTier"] == "Critical").sum()), "Actions": int(fdf["NeedsChange"].sum())}
    if not STATE["snapshots"] or STATE["snapshots"][-1] != snap:
        STATE["snapshots"].append(snap)
    STATE["snapshots"] = STATE["snapshots"][-20:]


def build_sankey_chart(fdf: pd.DataFrame) -> str:
    if fdf.empty:
        return no_data_chart("Sankey Lineage Visualization")
    sample = fdf.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(18)
    links: Dict[Tuple[str, str], float] = {}
    for _, r in sample.iterrows():
        src_table = str(r["NCH Target Table"])
        src_field = str(r["NCH Target Column"])
        tgt_field = str(r["SS Column"])
        report = str(r.get("Report", "")) or f"{r['Stream']} Quality Report"
        kpi = str(r.get("KPI", "")) or f"{r['Stream']} Match KPI"
        value = max(float(r.get("ImpactScore", 1) or 1), 0.35)
        for a, b in [(src_table, src_field), (src_field, tgt_field), (tgt_field, report), (report, kpi)]:
            links[(a, b)] = links.get((a, b), 0) + value
    nodes: List[str] = []
    for a, b in links:
        if a not in nodes:
            nodes.append(a)
        if b not in nodes:
            nodes.append(b)
    if not nodes or not links:
        return no_data_chart("Sankey Lineage Visualization", "No lineage relationships are available for the current selection.")
    idx = {n: i for i, n in enumerate(nodes)}
    fig = go.Figure(data=[go.Sankey(
        node=dict(label=nodes, pad=12, thickness=14, color="#d4af37", line=dict(color="#fff4b8", width=.5)),
        link=dict(source=[idx[a] for a, b in links], target=[idx[b] for a, b in links], value=[v for v in links.values()], color="rgba(249,217,118,.25)"),
    )])
    fig.update_layout(title="Sankey Lineage Visualization")
    return chart_html(fig, 480)


def build_dependency_chart(fdf: pd.DataFrame) -> str:
    if fdf.empty:
        return no_data_chart("Report Dependency Network")
    sample = fdf.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(8)
    nodes: List[str] = []
    edges: List[Tuple[str, str]] = []
    for _, r in sample.iterrows():
        source = str(r["NCH Target Table"])
        field = str(r["NCH Target Column"])
        report = str(r.get("Report", "")) or f"{r['Stream']} Quality Report"
        kpi = str(r.get("KPI", "")) or f"{r['Stream']} Match KPI"
        for n in [source, field, report, kpi]:
            if n not in nodes:
                nodes.append(n)
        edges.extend([(source, field), (field, report), (report, kpi)])
    if not nodes or not edges:
        return no_data_chart("Report Dependency Network", "No report dependencies are available for the current selection.")
    level_map = {n: (0 if "TABLE" in n else 1 if "FIELD" in n else 2 if "Report" in n or "Dashboard" in n else 3) for n in nodes}
    counts = {lvl: 0 for lvl in range(4)}
    totals = {lvl: sum(1 for n in nodes if level_map[n] == lvl) for lvl in range(4)}
    positions = {}
    for n in nodes:
        lvl = level_map[n]
        idx = counts[lvl]
        positions[n] = (lvl, (idx + 1) / (max(totals[lvl], 1) + 1))
        counts[lvl] += 1
    edge_x, edge_y = [], []
    for a, b in edges:
        x0, y0 = positions[a]
        x1, y1 = positions[b]
        edge_x += [x0, x1, None]
        edge_y += [y0, y1, None]
    xs, ys, text = [], [], []
    for n in nodes:
        x, y = positions[n]
        xs.append(x)
        ys.append(y)
        text.append(n)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=edge_x, y=edge_y, mode="lines", line=dict(width=1.2, color="#8a6a1f"), hoverinfo="none"))
    fig.add_trace(go.Scatter(x=xs, y=ys, mode="markers+text", text=text, textposition="top center", marker=dict(size=18, color="#d4af37", line=dict(width=2, color="#fff4b8"))))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    fig.update_layout(title="Report Dependency Network")
    return chart_html(fig, 420)


def build_charts(fdf: pd.DataFrame) -> Tuple[Dict[str, str], List[Dict[str, Any]]]:
    keys = ["hist", "stream", "weighted", "compare", "classmix", "risk", "heatmap", "trend", "drivers", "network", "sankey", "volume", "pipeline", "ownership", "impact"]
    charts = {k: "" for k in keys}
    compare_table: List[Dict[str, Any]] = []
    if fdf.empty:
        for key in keys:
            charts[key] = no_data_chart(key.title())
        return charts, []
    comparable = fdf[fdf["Comparable"]]
    if comparable.empty:
        charts["hist"] = no_data_chart("Comparable Field Match-Rate Distribution")
    else:
        fig = px.histogram(comparable, x="MatchRate", nbins=24, color="RiskTier", title="Comparable Field Match-Rate Distribution")
        fig.update_xaxes(tickformat=".0%")
        charts["hist"] = chart_html(fig)
    stream = fdf.groupby("Stream", as_index=False).agg(
        Fields=("MatchRate", "count"),
        AverageMatch=("MatchRate", "mean"),
        Matched=("MatchedClaims", "sum"),
        NotMatched=("NotMatchedClaims", "sum"),
        AverageImpact=("ImpactScore", "mean"),
        Critical=("RiskTier", lambda x: int((x == "Critical").sum())),
    )
    stream["WeightedMatch"] = stream["AverageMatch"]
    compare_table = stream.round(4).replace({np.nan: None}).to_dict("records")
    fig = px.bar(stream.sort_values("WeightedMatch"), x="Stream", y="WeightedMatch", text="WeightedMatch", title="Average Match Percentage by Stream")
    fig.update_yaxes(tickformat=".0%")
    fig.update_traces(texttemplate="%{text:.1%}", textposition="outside", marker_line_color="#f9e6a2", marker_line_width=1)
    charts["stream"] = chart_html(fig)

    weighted_volume = stream.copy()
    weighted_volume["TotalVolume"] = weighted_volume["Matched"].fillna(0) + weighted_volume["NotMatched"].fillna(0)
    if weighted_volume.empty or float(weighted_volume["TotalVolume"].sum()) <= 0:
        charts["weighted"] = no_data_chart(
            "Reconciliation Volume by Stream",
            "No matched or unmatched volume is available for the current selection.",
        )
    else:
        weighted_long = weighted_volume.melt(
            id_vars=["Stream"],
            value_vars=["Matched", "NotMatched"],
            var_name="Result",
            value_name="Volume",
        )
        weighted_fig = px.bar(
            weighted_long,
            x="Stream",
            y="Volume",
            color="Result",
            barmode="stack",
            title="Reconciliation Volume by Stream",
            category_orders={"Result": ["Matched", "NotMatched"]},
        )
        charts["weighted"] = chart_html(weighted_fig)

    if stream.empty or float(stream["NotMatched"].fillna(0).sum()) <= 0:
        charts["volume"] = no_data_chart(
            "Unmatched Volume by Stream",
            "No unmatched volume is available for the current selection.",
        )
    else:
        charts["volume"] = chart_html(
            px.bar(
                stream.sort_values("NotMatched", ascending=False),
                x="Stream",
                y="NotMatched",
                color="AverageImpact",
                title="Unmatched Volume by Stream",
                color_continuous_scale="YlOrBr",
            )
        )
    fig = px.scatter(stream, x="AverageMatch", y="WeightedMatch", size="Fields", color="Stream", hover_data=["Matched", "NotMatched", "AverageImpact"], title="Average Match Percentage by Stream")
    fig.update_xaxes(tickformat=".0%")
    fig.update_yaxes(tickformat=".0%")
    charts["compare"] = chart_html(fig)
    classmix = fdf.groupby(["Stream", "Classification"], as_index=False).size()
    charts["classmix"] = chart_html(px.bar(classmix, x="Stream", y="size", color="Classification", title="Classification Mix by Stream"))
    risk = fdf.groupby("RiskTier", as_index=False).size()
    risk = risk[risk["size"] > 0]
    if risk.empty:
        charts["risk"] = no_data_chart("Risk Distribution", "No risk-tier values are available for the current selection.")
    else:
        charts["risk"] = chart_html(px.pie(risk, values="size", names="RiskTier", hole=0.58, title="Risk Distribution"))
    heat = fdf.groupby(["Stream", "RiskTier"], as_index=False).agg(Impact=("ImpactScore", "mean"))
    charts["heatmap"] = chart_html(px.density_heatmap(heat, x="RiskTier", y="Stream", z="Impact", histfunc="avg", title="Average Impact Score Heatmap", color_continuous_scale="YlOrBr"))
    drivers = fdf.groupby("Sub-Classification", as_index=False).agg(Impact=("ImpactScore", "mean"), Fields=("ImpactScore", "count")).sort_values("Impact", ascending=False)
    charts["drivers"] = chart_html(px.bar(drivers, x="Sub-Classification", y="Impact", color="Fields", title="Top Risk Drivers", color_continuous_scale="YlOrBr"))
    add_snapshot(fdf)
    if len(STATE["snapshots"]) < 3:
        baseline = weighted_rate(fdf) if not pd.isna(weighted_rate(fdf)) else 0.85
        fa = field_average(fdf) if not pd.isna(field_average(fdf)) else 0.86
        snaps = pd.DataFrame({"Label": ["Baseline", "Current", "Projected"], "WeightedRate": [max(baseline - 0.035, 0), baseline, min(baseline + 0.018, 1)], "FieldAverage": [max(fa - 0.025, 0), fa, min(fa + 0.012, 1)]})
    else:
        snaps = pd.DataFrame(STATE["snapshots"])
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=snaps["Label"], y=snaps["WeightedRate"], mode="lines+markers", name="Weighted match"))
    fig.add_trace(go.Scatter(x=snaps["Label"], y=snaps["FieldAverage"], mode="lines+markers", name="Field average"))
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(title="Quality Trend")
    charts["trend"] = chart_html(fig)
    charts["network"] = build_dependency_chart(fdf)
    charts["sankey"] = build_sankey_chart(fdf)
    actions_count = int(fdf["NeedsChange"].sum())
    pipe = pd.DataFrame({"Stage": ["Detected", "Triaged", "Owner assigned", "In remediation", "Ready for validation"], "Items": [len(fdf), actions_count, max(int(actions_count * 0.65), 1), max(int(actions_count * 0.35), 1), max(int(actions_count * 0.18), 1)]})
    charts["pipeline"] = chart_html(px.bar(pipe, x="Stage", y="Items", title="Remediation Pipeline View", color="Items", color_continuous_scale="YlOrBr"))
    owner = pd.DataFrame({"Owner": ["Data Quality", "Claims Ops", "Reporting", "Engineering"], "Open Items": [max(int(actions_count*.34),1), max(int(actions_count*.27),1), max(int(actions_count*.22),1), max(int(actions_count*.17),1)]})
    charts["ownership"] = chart_html(px.pie(owner, names="Owner", values="Open Items", hole=0.45, title="Illustrative Open Items by Owner"))
    impact_rows = pd.DataFrame(STATE["impacts"] or [])
    if not impact_rows.empty:
        charts["impact"] = chart_html(px.bar(impact_rows, x="Report", color="Impact", title="Mapped Report Impact", barmode="group"), 380)
    else:
        charts["impact"] = no_data_chart("Mapped Report Impact")
    return charts, compare_table

# -----------------------------------------------------------------------------
# Frontend
# -----------------------------------------------------------------------------

HTML = r'''
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>DART - Data Assurance Reconciliation Tracker</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root{--bg:#100c05;--panel:#1b1408;--panel2:#261b0a;--panel3:#33250d;--line:#6f5420;--text:#fff7d1;--muted:#d8c891;--accent:#d4af37;--accent2:#f9d976;--accent3:#8a6a1f;--good:#9bd67d;--warn:#f9d976;--bad:#fb7185;--shadow:0 28px 78px rgba(0,0,0,.38)}
    *{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:radial-gradient(circle at 12% -8%,rgba(249,217,118,.18),transparent 32%),radial-gradient(circle at 85% 0%,rgba(212,175,55,.12),transparent 30%),linear-gradient(180deg,#100c05,#171006 45%,#0b0804);color:var(--text);font-family:Inter,ui-sans-serif,system-ui,Segoe UI,Arial,sans-serif}a{color:inherit;text-decoration:none}
    @keyframes pageFade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes cardIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}.page-enter{animation:pageFade .28s ease both}@media(prefers-reduced-motion:reduce){*,*:before,*:after{animation:none!important;transition:none!important}}
    .topbar{position:sticky;top:0;z-index:50;backdrop-filter:blur(18px);background:rgba(16,12,5,.93);border-bottom:1px solid rgba(249,217,118,.18)}.nav{display:flex;align-items:center;gap:16px;max-width:1600px;margin:auto;padding:13px 22px}.brand{display:flex;align-items:center;gap:10px;font-weight:950;white-space:nowrap}.gem{width:31px;height:31px;border-radius:10px;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);display:grid;place-items:center;color:#120e05;box-shadow:0 0 32px rgba(249,217,118,.25)}
    .navMenus{display:flex;gap:8px;flex-wrap:wrap;margin-left:auto}.menu{position:relative}.menuBtn{border:1px solid transparent;background:transparent;color:var(--muted);padding:9px 12px;border-radius:999px;cursor:pointer;font-size:.9rem;transition:background .18s ease,border-color .18s ease,color .18s ease,transform .18s ease}.menuBtn:hover,.menuBtn.active,.menu.open .menuBtn{color:#fff8cf;border-color:#d4af37;background:#3a2c10;transform:translateY(-1px)}.menuPanel{position:absolute;right:0;top:48px;background:rgba(27,20,8,.98);border:1px solid rgba(212,175,55,.42);border-radius:18px;box-shadow:0 28px 78px rgba(0,0,0,.45);min-width:255px;padding:10px;z-index:180;opacity:0;visibility:hidden;transform:translateY(-7px) scale(.985);pointer-events:none;transition:opacity .18s ease,transform .18s ease,visibility .18s ease}.menuPanel:before{content:"";position:absolute;left:0;right:0;top:-18px;height:18px}.menu.open .menuPanel{opacity:1;visibility:visible;transform:translateY(0) scale(1);pointer-events:auto}.menuPanel a{display:flex;align-items:center;gap:10px;padding:11px 12px;border-radius:12px;color:#decf91;cursor:pointer;transition:background .16s ease,color .16s ease,transform .16s ease}.menuPanel a:hover{background:#33250d;color:#fff7d1;transform:translateX(3px)}.wrap{max-width:1600px;margin:auto;padding:24px 22px 55px;animation:pageFade .28s ease both}
    .pageShell{display:grid;grid-template-columns:190px minmax(0,1fr);gap:18px;align-items:start}.sectionRail{position:sticky;top:86px;background:rgba(27,20,8,.94);border:1px solid rgba(212,175,55,.35);border-radius:20px;padding:13px;box-shadow:0 18px 50px rgba(0,0,0,.22)}.railTitle{font-size:.75rem;color:#f9e6a2;text-transform:uppercase;letter-spacing:.08em;margin:0 0 10px}.sectionRail a{display:block;color:#d8c891;border:1px solid transparent;border-radius:12px;padding:10px 10px;margin:4px 0;font-size:.86rem;transition:background .15s ease,transform .15s ease,color .15s ease}.sectionRail a:hover{background:#33250d;color:#fff7d1;transform:translateX(3px)}.sectionBlock{scroll-margin-top:92px;margin-bottom:16px}.sectionTitle h2{margin:0 0 10px}.hero{position:relative;overflow:hidden;border:1px solid var(--line);border-radius:30px;padding:34px;background:radial-gradient(circle at 10% 0%,rgba(249,217,118,.25),transparent 34%),radial-gradient(circle at 80% 0%,rgba(212,175,55,.18),transparent 34%),linear-gradient(135deg,#201606,#2b1e09,#120e05);box-shadow:var(--shadow);margin-bottom:18px}.hero h1{font-size:2.55rem;margin:0 0 8px}.hero p{color:var(--muted);max-width:980px;line-height:1.55}.heroActions{display:flex;gap:10px;flex-wrap:wrap;margin-top:20px}
    .panel{background:rgba(27,20,8,.96);border:1px solid rgba(212,175,55,.38);border-radius:22px;padding:18px;box-shadow:0 18px 54px rgba(0,0,0,.2);min-width:0;animation:cardIn .32s ease both;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}.panel:hover{transform:translateY(-2px);box-shadow:0 24px 64px rgba(0,0,0,.28);border-color:rgba(249,217,118,.48)}.panel h2,.panel h3{margin-top:0}.grid{display:grid;gap:16px}.grid6{grid-template-columns:repeat(6,minmax(0,1fr))}.grid4{grid-template-columns:repeat(4,minmax(0,1fr))}.grid3{grid-template-columns:repeat(3,minmax(0,1fr))}.grid2{grid-template-columns:repeat(2,minmax(0,1fr))}.miniDeck{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.miniCard{background:#130e05;border:1px solid rgba(212,175,55,.26);border-radius:16px;padding:14px}.miniCard b{display:block;font-size:1.18rem;margin-bottom:4px}.miniCard span{color:var(--muted);font-size:.84rem}
    .metric .label{color:var(--muted);font-size:.76rem;text-transform:uppercase;letter-spacing:.08em}.metric .value{font-size:1.85rem;font-weight:950;margin-top:7px}.metric .sub{color:var(--muted);font-size:.84rem}.health .value{font-size:2.8rem;background:linear-gradient(135deg,#fff4b8,#d4af37);-webkit-background-clip:text;color:transparent}
    .field{display:grid;gap:7px;min-width:0}.field label{font-size:.77rem;color:var(--muted);text-transform:uppercase;letter-spacing:.07em}input,select,textarea{background:#120d05;color:var(--text);border:1px solid rgba(212,175,55,.42);border-radius:13px;padding:12px 13px;outline:none;min-width:0;width:100%}input:focus,select:focus,textarea:focus{border-color:#f9d976;box-shadow:0 0 0 3px rgba(249,217,118,.12)}textarea{min-height:112px}.btn{border:0;border-radius:13px;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);padding:11px 15px;color:#130e05;font-weight:900;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:7px;transition:transform .16s ease,box-shadow .16s ease,filter .16s ease}.btn:hover{transform:translateY(-1px);box-shadow:0 12px 28px rgba(212,175,55,.18);filter:saturate(1.08)}.btn.secondary{background:#33250d;color:#fff7d1;border:1px solid #8a6a1f}.btn.ghost{background:transparent;color:#fff7d1;border:1px solid #8a6a1f}.btn.small{padding:8px 11px;font-size:.86rem}
    .filters summary{cursor:pointer;font-weight:900;color:#fff4b8}.filterGrid{display:grid;grid-template-columns:1.25fr repeat(4,minmax(145px,1fr));gap:12px;align-items:end;margin-top:15px}.advancedGrid{display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:12px;align-items:end;margin-top:14px}.filterPicker{background:#130e05;border:1px solid rgba(212,175,55,.3);border-radius:16px;padding:12px;min-height:91px}.filterPicker label{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.filterButton{width:100%;margin-top:8px;justify-content:space-between;background:#231906;color:#fff7d1;border:1px solid rgba(212,175,55,.42)}.filterCount{color:#f9d976;font-size:.82rem;font-weight:900}.activeFilters{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.filterTag{display:inline-flex;align-items:center;gap:8px;font-size:.82rem;border:1px solid rgba(212,175,55,.36);background:#130e05;color:#f6e7a4;border-radius:999px;padding:7px 10px}.filterTag button{border:0;background:transparent;color:#f9d976;cursor:pointer;font-weight:900}
    .modalBackdrop{position:fixed;inset:0;background:rgba(0,0,0,.58);backdrop-filter:blur(7px);display:none;align-items:center;justify-content:center;z-index:400;padding:22px}.modal{width:min(800px,96vw);max-height:86vh;overflow:auto;background:linear-gradient(180deg,#211707,#130e05);border:1px solid rgba(249,217,118,.45);border-radius:24px;box-shadow:0 32px 90px rgba(0,0,0,.55);padding:18px}.modalHeader{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;border-bottom:1px solid rgba(212,175,55,.22);padding-bottom:13px}.modalHeader h2{margin:0}.modalGrid{display:grid;grid-template-columns:1fr auto auto;gap:10px;margin:14px 0}.modalOptions{display:flex;gap:8px;flex-wrap:wrap;max-height:360px;overflow:auto;padding:10px;background:#100b04;border:1px solid rgba(212,175,55,.2);border-radius:16px}.chip{position:relative;display:inline-flex}.chip input{position:absolute;opacity:0;pointer-events:none}.chip span{border:1px solid rgba(212,175,55,.32);background:#231906;color:#decf91;border-radius:999px;padding:8px 11px;font-size:.86rem;cursor:pointer;user-select:none}.chip input:checked + span{background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);color:#130e05;border-color:#f9d976;font-weight:900}.rangeLine{font-size:.83rem;color:var(--muted);margin-top:4px}
    .tableWrap{width:100%;overflow:auto;border-radius:16px;border:1px solid rgba(212,175,55,.24)}.table{width:100%;border-collapse:collapse;min-width:1050px}.table th{position:sticky;top:0;font-size:.75rem;color:#f9e6a2;text-transform:uppercase;letter-spacing:.06em;text-align:left;background:#160f05;z-index:1}.table th,.table td{padding:11px 12px;border-bottom:1px solid rgba(212,175,55,.18);vertical-align:top}.table tr:hover td{background:rgba(249,217,118,.045)}.pill{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border-radius:999px;background:#231906;border:1px solid #8a6a1f;color:#f6e7a4;font-size:.78rem}.pill.Critical{border-color:rgba(251,113,133,.7);color:#fecdd3}.pill.Elevated{border-color:#f9d976;color:#f9e6a2}.pill.Watch{border-color:#d4af37;color:#ffe8a3}.pill.Stable{border-color:rgba(155,214,125,.55);color:#d7f8cd}.bar{height:9px;border-radius:999px;background:#3a2c10;overflow:hidden;min-width:100px}.bar span{display:block;height:100%;background:linear-gradient(90deg,#fb7185,#f9d976,#9bd67d)}.insight{border-left:4px solid #d4af37;padding:15px 16px;background:#171006;border-radius:13px;color:#fff7d1}.muted{color:var(--muted)}.chat{display:grid;gap:12px;max-height:520px;overflow:auto}.bubble{padding:14px 16px;border-radius:16px;max-width:900px;white-space:pre-wrap}.user{background:#3a2c10;margin-left:auto}.assistant{background:#171006;border:1px solid #6f5420}.footer{margin-top:25px;color:var(--muted);font-size:.82rem}.kanban{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:14px}.lane{background:#130e05;border:1px solid rgba(212,175,55,.32);border-radius:18px;padding:13px;min-height:240px}.lane h4{margin:0 0 10px}.task{background:#231906;border:1px solid #6f5420;border-radius:15px;padding:12px;margin-bottom:10px}.empty{padding:28px;border:1px dashed #8a6a1f;border-radius:16px;color:var(--muted);text-align:center;background:#130e05}.toast{position:fixed;bottom:22px;right:22px;background:#2b1e09;border:1px solid #d4af37;border-radius:14px;padding:13px 16px;box-shadow:var(--shadow);z-index:100;display:none}.briefText{white-space:pre-wrap;line-height:1.55;background:#130e05;border:1px solid rgba(212,175,55,.24);padding:16px;border-radius:16px;color:#fff7d1}.scorecard h2{font-size:1.5rem;margin-bottom:4px}.scoreLine{display:flex;justify-content:space-between;border-top:1px solid rgba(212,175,55,.16);padding:9px 0;color:#decf91}.navNote{font-size:.78rem;color:#d8c891;margin-left:4px}.explain{line-height:1.55}.explain b{color:#fff4b8}.editRow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px}
    @media(max-width:1180px){.pageShell,.grid6,.grid4,.grid3,.grid2,.kanban,.miniDeck,.filterGrid,.advancedGrid,.modalGrid,.editRow{grid-template-columns:1fr}.sectionRail{position:relative;top:auto}.nav{align-items:flex-start;flex-direction:column}.navMenus{margin-left:0}.hero h1{font-size:1.9rem}.menuPanel{position:static;display:none;opacity:1;visibility:visible;transform:none;pointer-events:auto}.menu.open .menuPanel{display:block}}
  </style>
</head>
<body>
  <div class="topbar"><div class="nav"><a class="brand" href="#" onclick="showPage('home')"><span class="gem">◆</span><span>DART</span></a><div class="navMenus" id="navlinks"></div></div></div>
  <main class="wrap"><section id="app"></section><div class="footer" id="footer"></div></main><div class="toast" id="toast"></div>
  <div class="modalBackdrop" id="filterModal"><div class="modal"><div class="modalHeader"><div><h2 id="modalTitle">Filter</h2><p class="muted" id="modalSub">Choose values to include in the current scope.</p></div><button class="btn ghost" onclick="closeFilterModal()">Close</button></div><div class="modalGrid"><input id="modalSearch" placeholder="Search values" oninput="renderModalOptions()"><button class="btn secondary" onclick="modalSelectAll(true)">Select visible</button><button class="btn ghost" onclick="modalSelectAll(false)">Clear visible</button></div><div class="modalOptions" id="modalOptions"></div><br><div class="heroActions"><button class="btn" onclick="applyFilterModal()">Apply filter</button><button class="btn secondary" onclick="clearCurrentFilter()">Clear this filter</button></div></div></div>
<script>
const navGroups=[['Overview',[['home','Command Center'],['briefing','Briefing'],['metrics','Metric Explanations'],['briefbuilder','Executive Brief']]],['Intelligence',[['riskcenter','Risk Center'],['insights','AI Insights'],['scorecards','Scorecards'],['simulator','Impact Simulator']]],['Operations',[['actioncenter','Remediation Center'],['impactexplorer','Impact Explorer'],['lineage','Data Lineage'],['catalog','Mapping Catalog'],['governance','Governance Center']]],['Data',[['explorer','System Explorer'],['quality','Quality Analytics'],['compare','Stream Comparison'],['data','Data Management']]],['More',[['copilot','DART Copilot'],['profile','Profile'],['settings','Settings']]]];
const filterPages=new Set(['home','briefing','explorer','quality','compare','lineage','catalog','impactexplorer','briefbuilder','riskcenter','insights','scorecards','simulator','actioncenter','governance']);
let state={page:'home',filters:{streams:null,classes:null,tiers:null,reasons:null,min_rate:0,min_impact:0,min_unmatched:0,actions_only:false,search:'',custom_filters:[]},meta:null};
let modalState={key:null,title:'',items:[],selected:[]};
function esc(v){return String(v??'').replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));}
function pctNum(v){return v===null||v===undefined||Number.isNaN(Number(v))?'-':(Number(v)*100).toFixed(1)+'%';}
function intFmt(v){return Number(v||0).toLocaleString();}
async function api(path,opts={}){const r=await fetch(path,opts);if(!r.ok)throw new Error(await r.text());return await r.json();}
async function postJson(path,body){return api(path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});}
async function putJson(path,body){return api(path,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});}
async function del(path){return api(path,{method:'DELETE'});}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.style.display='block';setTimeout(()=>t.style.display='none',2200)}
function pageInGroup(group){return group[1].some(([id])=>id===state.page)}
function closeAllMenus(){document.querySelectorAll('.menu').forEach(m=>{m.classList.remove('open');const b=m.querySelector('.menuBtn');if(b)b.setAttribute('aria-expanded','false');});}
function toggleMenu(e,menu){e.stopPropagation();const wasOpen=menu.classList.contains('open');closeAllMenus();if(!wasOpen){menu.classList.add('open');const b=menu.querySelector('.menuBtn');if(b)b.setAttribute('aria-expanded','true');}}
document.addEventListener('click',closeAllMenus);document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeAllMenus();closeFilterModal();}});
function setNav(){document.getElementById('navlinks').innerHTML=navGroups.map(g=>`<div class="menu ${pageInGroup(g)?'current':''}"><button class="menuBtn ${pageInGroup(g)?'active':''}" onclick="toggleMenu(event,this.parentElement)" aria-haspopup="true" aria-expanded="false">${g[0]} ▾</button><div class="menuPanel">${g[1].map(([id,label])=>`<a onclick="showPage('${id}');closeAllMenus()">${label}</a>`).join('')}</div></div>`).join('');}
async function boot(){state.meta=await api('/api/meta');const target=sessionStorage.getItem('dart_target_page');if(target){state.page=target;sessionStorage.removeItem('dart_target_page');}if(!state.meta.auth_user){renderAuth();return;}if(!state.meta.persona){renderOnboarding();return;}setNav();await render();}
function sectionShell(items,content){return `<div class="pageShell"><aside class="sectionRail"><div class="railTitle">On this page</div>${items.map(x=>`<a href="#${x[0]}">${x[1]}</a>`).join('')}</aside><div>${content}</div></div>`;}
function section(id,title,content){return `<section class="sectionBlock" id="${id}"><div class="sectionTitle"><h2>${title}</h2></div>${content}</section>`;}
function renderAuth(){document.getElementById('navlinks').innerHTML='';document.getElementById('footer').innerHTML='';document.getElementById('app').innerHTML=`<div class="hero"><h1>DART</h1><p><b>Data Assurance Reconciliation Tracker</b>. Sign in or create a local prototype account to save a profile in <code>dart_users.json</code>.</p></div><div class="grid grid2"><form class="panel" id="loginForm"><h2>Login</h2><div class="field"><label>Username</label><input id="loginUser" autocomplete="username"></div><div class="field"><label>Password</label><input id="loginPass" type="password" autocomplete="current-password"></div><br><button class="btn" type="submit">Login</button></form><form class="panel" id="signupForm"><h2>Create profile</h2><p class="muted">Saved to <code>dart_users.json</code> in the same folder as this app.</p><div class="field"><label>Username</label><input id="signupUser" autocomplete="username"></div><div class="field"><label>Password</label><input id="signupPass" type="password" autocomplete="new-password"></div><div class="field"><label>Display name</label><input id="signupName" placeholder="Nick Holmes"></div><div class="field"><label>Email</label><input id="signupEmail" placeholder="name@example.com"></div><div class="field"><label>Organization</label><input id="signupOrg" placeholder="Team or organization"></div><br><button class="btn" type="submit">Create profile</button></form></div>`;document.getElementById('loginForm').onsubmit=submitLogin;document.getElementById('signupForm').onsubmit=submitSignup;}
async function submitLogin(e){e.preventDefault();try{await postJson('/api/login',{username:loginUser.value,password:loginPass.value});toast('Logged in');window.location.reload();}catch(err){console.error(err);toast('Login failed');}}
async function submitSignup(e){e.preventDefault();try{await postJson('/api/signup',{username:signupUser.value,password:signupPass.value,display_name:signupName.value,email:signupEmail.value,organization:signupOrg.value});toast('Profile created');window.location.reload();}catch(err){console.error(err);toast('Could not create profile');}}
async function logoutUser(){await postJson('/api/logout',{});toast('Logged out');window.location.reload();}
async function saveProfile(e){if(e)e.preventDefault();await postJson('/api/profile',{display_name:profileName.value,email:profileEmail.value,organization:profileOrg.value,role:profileRole.value,password:profilePass.value});toast('Profile saved');state.meta=await api('/api/meta');await render();}
function renderOnboarding(){document.getElementById('navlinks').innerHTML='';document.getElementById('footer').innerHTML='';document.getElementById('app').innerHTML=`<div class="hero"><h1>Build your DART workspace</h1><p>Choose your lens once, then DART opens with command center, metric explanations, editable lineage, impact explorer, seeded issues, and governance center.</p></div><form class="panel" id="personaForm"><h2>Personalize the experience</h2><div class="grid grid2"><div class="field"><label>Program</label><select id="program"><option>Medicare</option><option>Medicaid</option><option>Both</option><option>Other / General</option></select></div><div class="field"><label>Primary role</label><select id="role"><option>Executive / Leadership</option><option>Data / Analytics</option><option>Program / Policy</option><option>Operations</option><option>Quality / Compliance</option><option>IT / Engineering</option><option>Research</option><option>Other</option></select></div><div class="field"><label>Focus areas</label><input id="focus" value="Data quality, Claims, Reporting impact, AI insights"></div><div class="field"><label>Audience</label><select id="audience"><option>Leadership</option><option>Myself</option><option>Analysts</option><option>Program teams</option><option>Technical teams</option><option>External stakeholders</option></select></div><div class="field"><label>Detail level</label><select id="depth"><option>Executive</option><option selected>Balanced</option><option>Technical</option></select></div><div class="field"><label>First question</label><input id="first_question" placeholder="Which areas need the most attention?"></div></div><br><button class="btn" type="button" onclick="savePersona(event)">Build Workspace</button></form>`;personaForm.onsubmit=savePersona;}
async function savePersona(e){if(e)e.preventDefault();const payload={program:program.value,role:role.value,audience:audience.value,depth:depth.value,focus:focus.value.split(',').map(x=>x.trim()).filter(Boolean),geography:['National'],first_question:first_question.value};await postJson('/api/persona',payload);sessionStorage.setItem('dart_target_page','home');toast('Workspace created');window.location.reload();}
async function saveSettings(e){if(e)e.preventDefault();await postJson('/api/persona',{program:setProgram.value,role:setRole.value,audience:setAudience.value,depth:setDepth.value,focus:setFocus.value.split(',').map(x=>x.trim()).filter(Boolean),geography:['National'],first_question:setQuestion.value});toast('Settings updated');state.meta=await api('/api/meta');await render();}
async function showPage(id){state.page=id;setNav();await render();}
function filterLabel(key){const arr=state.filters[key];if(arr===null||arr===undefined)return 'All values';if(!arr.length)return 'No values';return arr.length===1?arr[0]:`${arr.length} selected`;}
function openFilterModal(key,title,items){modalState={key,title,items:[...items],selected:[...(state.filters[key]||items)]};modalTitle.textContent=title;modalSub.textContent='Search and select values. Leave everything selected to include all values.';modalSearch.value='';filterModal.style.display='flex';renderModalOptions();}
function closeFilterModal(){filterModal.style.display='none';}
function visibleModalItems(){const q=modalSearch.value.toLowerCase().trim();return modalState.items.filter(x=>String(x).toLowerCase().includes(q));}
function safeArg(v){return String(v).replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
function renderModalOptions(){const visible=visibleModalItems();modalOptions.innerHTML=visible.map(x=>`<label class="chip"><input type="checkbox" data-val="${esc(x)}" ${modalState.selected.includes(x)?'checked':''} onchange="toggleModalValue('${safeArg(x)}',this.checked)"><span>${esc(x)}</span></label>`).join('')||'<div class="empty">No matching values.</div>';}
function toggleModalValue(v,on){if(on&&!modalState.selected.includes(v))modalState.selected.push(v);if(!on)modalState.selected=modalState.selected.filter(x=>x!==v);}
function modalSelectAll(on){const visible=visibleModalItems();if(on){visible.forEach(v=>{if(!modalState.selected.includes(v))modalState.selected.push(v);});}else{modalState.selected=modalState.selected.filter(v=>!visible.includes(v));}renderModalOptions();}
function applyFilterModal(){const all=modalState.items;const sel=modalState.selected;if(sel.length===0){state.filters[modalState.key]=[];}else if(sel.length===all.length){state.filters[modalState.key]=null;}else{state.filters[modalState.key]=sel;}closeFilterModal();render();}
function clearCurrentFilter(){state.filters[modalState.key]=null;closeFilterModal();render();}
function removeCustomFilter(i){state.filters.custom_filters.splice(i,1);render();}
function addCustomFilter(){const field=customField.value;const op=customOp.value;const value=customValue.value.trim();if(!field||!value){toast('Choose a field and value');return;}state.filters.custom_filters=state.filters.custom_filters||[];state.filters.custom_filters.push({field,op,value});customValue.value='';render();}
function activeFilterTags(){const tags=[];[['streams','Streams'],['classes','Classification'],['tiers','Risk tier'],['reasons','Drivers']].forEach(([k,l])=>{const v=state.filters[k];if(v&&v.length)tags.push(`<span class="filterTag">${l}: ${esc(v.join(', '))}<button onclick="state.filters.${k}=null;render()">x</button></span>`);});(state.filters.custom_filters||[]).forEach((f,i)=>tags.push(`<span class="filterTag">${esc(f.field)} ${esc(f.op)} ${esc(f.value)}<button onclick="removeCustomFilter(${i})">x</button></span>`));return tags.length?`<div class="activeFilters">${tags.join('')}</div>`:'';}
function filterBar(meta){if(!filterPages.has(state.page))return '';const cols=(meta.columns_list||[]).map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');return `<details class="panel filters" style="margin-bottom:16px"><summary>Data scope and filters <span class="navNote">Click a filter to open a searchable selector. Add custom field filters below.</span></summary><div class="filterGrid"><div class="field"><label>Search everything</label><input id="search" placeholder="Table, field, stream, system, recommendation" value="${esc(state.filters.search)}"></div><div class="filterPicker"><label>Streams</label><button class="btn filterButton" onclick='openFilterModal("streams","Streams",${JSON.stringify(meta.streams)})'><span>${esc(filterLabel('streams'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Classification</label><button class="btn filterButton" onclick='openFilterModal("classes","Classification",${JSON.stringify(meta.classes)})'><span>${esc(filterLabel('classes'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Risk tier</label><button class="btn filterButton" onclick='openFilterModal("tiers","Risk tier",${JSON.stringify(meta.tiers)})'><span>${esc(filterLabel('tiers'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Drivers</label><button class="btn filterButton" onclick='openFilterModal("reasons","Drivers",${JSON.stringify(meta.reasons)})'><span>${esc(filterLabel('reasons'))}</span><span class="filterCount">Choose values ▾</span></button></div></div><div class="advancedGrid"><div class="field"><label>Custom field</label><select id="customField">${cols}</select></div><div class="field"><label>Operator</label><select id="customOp"><option value="contains">contains</option><option value="equals">equals</option><option value="not_equals">does not equal</option><option value="starts_with">starts with</option><option value="ends_with">ends with</option><option value="gt">greater than</option><option value="gte">greater than or equal</option><option value="lt">less than</option><option value="lte">less than or equal</option></select></div><div class="field"><label>Value</label><input id="customValue" placeholder="Value to filter on"></div><button class="btn" onclick="addCustomFilter()">Add custom filter</button></div><div class="advancedGrid"><div class="field"><label>Min match %</label><input id="minRate" type="range" min="0" max="100" value="${state.filters.min_rate*100}" oninput="minRateLabel.textContent=this.value+'%'"><div class="rangeLine">Current: <span id="minRateLabel">${state.filters.min_rate*100}%</span></div></div><div class="field"><label>Min impact</label><input id="minImpact" type="number" min="0" value="${state.filters.min_impact||0}"></div><div class="field"><label>Min unmatched</label><input id="minUnmatched" type="number" min="0" value="${state.filters.min_unmatched||0}"></div><div class="field"><label>View</label><select id="actionsOnly"><option value="false">All rows</option><option value="true" ${state.filters.actions_only?'selected':''}>Needs action only</option></select></div></div><br><button class="btn" onclick="readFilters();render()">Apply filters</button> <button class="btn secondary" onclick="resetFilters()">Reset filters</button>${activeFilterTags()}</details>`;}
function readFilters(){state.filters.min_rate=Number(minRate.value||0)/100;state.filters.min_impact=Number(minImpact.value||0);state.filters.min_unmatched=Number(minUnmatched.value||0);state.filters.actions_only=actionsOnly.value==='true';state.filters.search=search.value;state.filters.custom_filters=state.filters.custom_filters||[];}
async function resetFilters(){state.filters={streams:null,classes:null,tiers:null,reasons:null,min_rate:0,min_impact:0,min_unmatched:0,actions_only:false,search:'',custom_filters:[]};await render();}
async function getData(){return postJson('/api/view',state.filters);}
function metrics(s){return `<div class="grid grid6"><div class="panel metric health"><div class="label">Health score</div><div class="value">${s.health_score}</div><div class="sub">Event-weighted quality</div></div><div class="panel metric"><div class="label">Risk score</div><div class="value">${s.risk_score}</div><div class="sub">Avg impact score</div></div><div class="panel metric"><div class="label">Critical items</div><div class="value">${s.critical}</div><div class="sub">Highest risk tier</div></div><div class="panel metric"><div class="label">Open actions</div><div class="value">${s.actions}</div><div class="sub">Needs review/change</div></div><div class="panel metric"><div class="label">Unmatched volume</div><div class="value">${s.unmatched}</div><div class="sub">Current scope</div></div><div class="panel metric"><div class="label">Affected reports</div><div class="value">${s.affected_reports}</div><div class="sub">Mapped or estimated</div></div></div>`;}
function miniDeck(s){return `<div class="miniDeck"><div class="miniCard"><b>${s.weighted_rate}</b><span>Weighted quality signal</span></div><div class="miniCard"><b>${s.risk_count}</b><span>Elevated or critical risk fields</span></div><div class="miniCard"><b>${s.affected_reports}</b><span>Reports that may need review</span></div></div>`;}
function table(rows,limitCols=false){const cols=limitCols?['Stream','NCH Target Column','MatchRate','NotMatchedClaims','RiskTier','ImpactScore','Recommendation']:['Stream','NCH Target Table','NCH Target Column','SS Table','SS Column','MatchedClaims','NotMatchedClaims','MatchRate','Classification','RiskTier','ImpactScore','Sub-Classification','Disposition','Recommendation'];return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>{let v=r[c];if(c==='MatchRate')return `<td><div>${pctNum(v)}</div><div class="bar"><span style="width:${Math.max(0,Math.min(100,Number(v||0)*100))}%"></span></div></td>`;if(c==='RiskTier')return `<td><span class="pill ${esc(v)}">${esc(v)}</span></td>`;return `<td>${esc(v)}</td>`}).join('')}</tr>`).join('')}</tbody></table></div>`;}
function simpleTable(rows){if(!rows||!rows.length)return '<div class="empty">Nothing saved yet. Starter examples will appear where helpful.</div>';const cols=Object.keys(rows[0]);return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;}
function cards(items){return items.map(r=>`<div class="insight"><b>${esc(r.Stream)} · ${esc(r['NCH Target Column'])}</b><br><span class="muted">${pctNum(r.MatchRate)} match · ${intFmt(r.NotMatchedClaims)} unmatched · ${esc(r.RiskTier)} risk · impact ${esc(r.ImpactScore)}</span><br>${esc(r.Recommendation)}</div>`).join('<br>')||'<div class="empty">No priority findings in the current scope.</div>';}
function runInjectedScripts(root){root.querySelectorAll('script').forEach(oldScript=>{const newScript=document.createElement('script');Array.from(oldScript.attributes).forEach(attr=>newScript.setAttribute(attr.name,attr.value));newScript.textContent=oldScript.textContent;oldScript.parentNode.replaceChild(newScript,oldScript);});if(window.Plotly){setTimeout(()=>{root.querySelectorAll('.plotly-graph-div').forEach(g=>{try{Plotly.Plots.resize(g);}catch(err){console.warn('Plotly resize skipped',err);}});},75);}}
function executiveBrief(data,meta){const s=data.summary;const p=meta.persona||{};const top=(s.top_actions||[]).slice(0,6).map((r,i)=>`${i+1}. ${r.Stream} · ${r['NCH Target Column']} | ${pctNum(r.MatchRate)} match | ${intFmt(r.NotMatchedClaims)} unmatched | ${r.RiskTier} | impact ${r.ImpactScore}`).join('\n')||'No priority findings in the current scope.';return `Executive Brief\n\nAudience: ${p.audience||'Leadership'}\nProgram: ${p.program||'Not specified'}\nRole lens: ${p.role||'Not specified'}\n\nExecutive Risk Rating\n- Health score: ${s.health_score}\n- Risk score: ${s.risk_score}\n- Critical items: ${s.critical}\n- Open actions: ${s.actions}\n- Affected reports: ${s.affected_reports}\n\nOverall Data Quality\n- Field average: ${s.field_average}\n- Event-weighted match: ${s.weighted_rate}\n- Unmatched volume: ${s.unmatched}\n\nPriority Findings\n${top}\n\nDecision Support Recommendations\n1. Validate highest-impact mappings first.\n2. Confirm lineage for affected fields and downstream reports.\n3. Assign owners for critical and elevated items in the Remediation Center.\n4. Use Impact Explorer for fields that support executive KPIs.\n5. Review Governance Center controls for ownership and evidence.\n\nDecision Required\nConfirm whether critical mappings should be prioritized for remediation in the next review cycle.`}
function aiInsights(data){const s=data.summary;const t=s.top_actions||[];const top=t[0]||{};const second=t[1]||{};const stream=(s.stream_summary||[])[0]||{};return [{title:'Highest priority risk',body:top.Stream?`${top.Stream} · ${top['NCH Target Column']} has ${pctNum(top.MatchRate)} match with ${intFmt(top.NotMatchedClaims)} unmatched claims.`:'No priority risk in current scope.',action:'Validate mapping and lineage for this field.'},{title:'Largest contributor by stream',body:stream.Stream?`${stream.Stream} contributes the largest unmatched volume in the current scope.`:'No stream summary available.',action:'Review stream-level scorecard and issue queue.'},{title:'Next best action',body:second.Stream?`After the top risk, review ${second.Stream} · ${second['NCH Target Column']} because it also appears in the priority queue.`:'Create an issue only if a priority finding appears.',action:'Use the Remediation Center to assign an owner and status.'},{title:'Governance angle',body:`Current health score is ${s.health_score}, with ${s.actions} action candidates and ${s.affected_reports} affected reports.`,action:'Use Governance Center to capture owners, controls, and evidence.'}]}
function insightsHtml(items){return `<div class="grid grid4">${items.map(x=>`<div class="panel"><h3>${esc(x.title)}</h3><p>${esc(x.body)}</p><p class="muted"><b>Recommendation:</b> ${esc(x.action)}</p></div>`).join('')}</div>`}
function scorecards(rows){return `<div class="grid grid3">${rows.map(r=>`<div class="panel scorecard"><h2>${esc(r.Stream)}</h2><p class="muted">Stream scorecard</p><div class="scoreLine"><span>Overall average match</span><b>${pctNum(r.WeightedMatch)}</b></div><div class="scoreLine"><span>Average match</span><b>${pctNum(r.AverageMatch)}</b></div><div class="scoreLine"><span>Fields</span><b>${intFmt(r.Fields)}</b></div><div class="scoreLine"><span>Not matched</span><b>${intFmt(r.NotMatched)}</b></div><div class="scoreLine"><span>Critical</span><b>${intFmt(r.Critical)}</b></div><div class="scoreLine"><span>Avg impact</span><b>${Number(r.AverageImpact||0).toFixed(2)}</b></div></div>`).join('')}</div>`}
function metricExplanations(){return `<div class="grid grid3 explain"><div class="panel"><h3>Health score</h3><p><b>What it means:</b> Event-weighted match quality across the current filtered scope.</p><p><b>How to use it:</b> Use it as the executive quality signal because higher-volume fields influence it more.</p></div><div class="panel"><h3>Field average</h3><p><b>What it means:</b> Average match rate across comparable fields, treating each field equally.</p><p><b>How to use it:</b> Good for breadth. It can hide volume concentration, so compare it with weighted match.</p></div><div class="panel"><h3>Weighted rate</h3><p><b>What it means:</b> Matched claims divided by total claims in scope.</p><p><b>How to use it:</b> Good for workload and business-volume weighting.</p></div><div class="panel"><h3>Risk score</h3><p><b>What it means:</b> Average impact score in the scope.</p><p><b>Formula:</b> (1 - Match Rate) x log(1 + Total Claims).</p></div><div class="panel"><h3>Risk tier</h3><p><b>What it means:</b> Stable, Watch, Elevated, or Critical based on impact score.</p><p><b>How to use it:</b> Sort remediation work by tier before drilling into drivers.</p></div><div class="panel"><h3>Affected reports</h3><p><b>What it means:</b> Reports/KPIs mapped in Impact Explorer or inferred from the current stream/report scope.</p><p><b>How to use it:</b> Converts technical variance into stakeholder impact.</p></div></div>`;}
function lineageEditor(rows){return `<div class="panel"><h3>Add or update mapping</h3><div class="editRow"><input id="linId" placeholder="ID, leave blank for new"><input id="linStream" placeholder="Stream"><input id="linSourceTable" placeholder="Source table"><input id="linSourceField" placeholder="Source field"></div><div class="editRow"><input id="linTargetTable" placeholder="Target table"><input id="linTargetField" placeholder="Target field"><input id="linReport" placeholder="Report"><input id="linKPI" placeholder="KPI"></div><div class="editRow"><input id="linOwner" placeholder="Owner"><select id="linStatus"><option>Draft</option><option>Validated</option><option>Needs review</option><option>Retired</option></select><input id="linRiskTier" placeholder="Risk tier"><input id="linNotes" placeholder="Notes"></div><button class="btn" onclick="saveLineage()">Save mapping</button> <button class="btn secondary" onclick="seedLineage()">Seed from current data</button></div><br><div class="panel"><h3>Active mappings</h3>${editableLineageTable(rows)}</div>`;}
function editableLineageTable(rows){if(!rows||!rows.length)return '<div class="empty">No mappings yet.</div>';const cols=['id','Stream','SourceTable','SourceField','TargetTable','TargetField','Report','KPI','Owner','Status','RiskTier','Notes'];return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}<th>Actions</th></tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}<td><button class="btn small secondary" onclick='loadLineage(${JSON.stringify(r).replace(/'/g,"'")})'>Edit</button> <button class="btn small ghost" onclick="deleteLineage('${esc(r.id)}')">Delete</button></td></tr>`).join('')}</tbody></table></div>`;}
function loadLineage(r){linId.value=r.id||'';linStream.value=r.Stream||'';linSourceTable.value=r.SourceTable||'';linSourceField.value=r.SourceField||'';linTargetTable.value=r.TargetTable||'';linTargetField.value=r.TargetField||'';linReport.value=r.Report||'';linKPI.value=r.KPI||'';linOwner.value=r.Owner||'';linStatus.value=r.Status||'Draft';linRiskTier.value=r.RiskTier||'';linNotes.value=r.Notes||'';window.scrollTo({top:0,behavior:'smooth'});}
async function saveLineage(){const payload={id:linId.value,Stream:linStream.value,SourceTable:linSourceTable.value,SourceField:linSourceField.value,TargetTable:linTargetTable.value,TargetField:linTargetField.value,Report:linReport.value,KPI:linKPI.value,Owner:linOwner.value,Status:linStatus.value,RiskTier:linRiskTier.value,Notes:linNotes.value};await postJson('/api/lineage',payload);toast('Mapping saved');await render();}
async function deleteLineage(id){await del('/api/lineage/'+encodeURIComponent(id));toast('Mapping deleted');await render();}
async function seedLineage(){await postJson('/api/seed-lineage',{});toast('Lineage seeded from current data');await render();}
function impactEditor(rows, dataRows){return `<div class="grid grid2"><div class="panel"><h3>Map a field to a report/KPI</h3><div class="field"><label>Field</label><select id="impactField">${dataRows.map((r,i)=>`<option value="${i}">${esc(r.Stream)} · ${esc(r['NCH Target Column'])}</option>`).join('')}</select></div><div class="field"><label>Report / dashboard</label><input id="impactReport" placeholder="Claims Integrity Dashboard"></div><div class="field"><label>KPI</label><input id="impactKpi" placeholder="Match Rate, Paid Claims, Denials"></div><div class="field"><label>Business owner</label><input id="impactOwner" placeholder="Reporting"></div><div class="field"><label>Impact</label><select id="impactLevel"><option>Low</option><option>Moderate</option><option>High</option><option>Critical</option></select></div><div class="field"><label>Decision need</label><input id="impactDecision" placeholder="Confirm remediation priority"></div><br><button class="btn" onclick='addImpact(${JSON.stringify(dataRows).replace(/'/g,"'")})'>Map report impact</button></div><div class="panel"><h3>Saved impact map</h3>${simpleTable(rows)}</div></div>`;}
function governanceEditor(rows){return `<div class="grid grid2"><div class="panel"><h3>Create governance control</h3><input id="govId" placeholder="ID, leave blank for new"><br><br><input id="govControl" placeholder="Control"><br><br><input id="govOwner" placeholder="Owner"><br><br><input id="govCadence" placeholder="Cadence"><br><br><select id="govStatus"><option>Draft</option><option>Planned</option><option>Active</option><option>Blocked</option><option>Retired</option></select><br><br><textarea id="govEvidence" placeholder="Evidence source or notes"></textarea><br><button class="btn" onclick="saveGovernance()">Save control</button></div><div class="panel"><h3>Governance controls</h3>${governanceTable(rows)}</div></div>`;}
function governanceTable(rows){if(!rows||!rows.length)return '<div class="empty">No governance controls saved.</div>';const cols=['id','Control','Owner','Cadence','Status','Evidence'];return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}<th>Actions</th></tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}<td><button class="btn small secondary" onclick='loadGovernance(${JSON.stringify(r).replace(/'/g,"'")})'>Edit</button> <button class="btn small ghost" onclick="deleteGovernance('${esc(r.id)}')">Delete</button></td></tr>`).join('')}</tbody></table></div>`;}
function loadGovernance(r){govId.value=r.id||'';govControl.value=r.Control||'';govOwner.value=r.Owner||'';govCadence.value=r.Cadence||'';govStatus.value=r.Status||'Draft';govEvidence.value=r.Evidence||'';}
async function saveGovernance(){await postJson('/api/governance',{id:govId.value,Control:govControl.value,Owner:govOwner.value,Cadence:govCadence.value,Status:govStatus.value,Evidence:govEvidence.value});toast('Governance control saved');await render();}
async function deleteGovernance(id){await del('/api/governance/'+encodeURIComponent(id));toast('Governance control deleted');await render();}
function kanban(rows){const lanes=['Open','In progress','Blocked','Resolved'];return `<div class="kanban">${lanes.map(l=>`<div class="lane"><h4>${l}</h4>${(rows||[]).filter(r=>r.Status===l).map(r=>`<div class="task"><b>${esc(r.Field)}</b><br><span class="muted">${esc(r.Stream)} · ${esc(r.Priority||'')}</span><br>${esc(r.Owner||'Unassigned')}<br><span class="muted">${esc(r.Note||'')}</span></div>`).join('')||'<div class="muted">No items</div>'}</div>`).join('')}</div>`;}
async function render(){const meta=await api('/api/meta');state.meta=meta;const data=await getData();document.getElementById('footer').innerHTML=`Loaded: ${esc(meta.source)} · ${intFmt(meta.rows)} rows · ${intFmt(meta.columns)} columns · DART ${esc(meta.version)} one-file FastAPI command center`;const p=meta.persona||{};let html=filterBar(meta);
if(state.page==='home'){let body=`${section('overview','Overview',`<div class="hero"><h1>DART Command Center</h1><p>A consolidated view of data quality, reconciliation risk, unmatched volume, remediation activity, and downstream business impact${p.program||p.role?` for <b>${esc(p.program||'your program')}</b> with a <b>${esc(p.role||'general')}</b> role lens`:''}. Use this page as your primary operational command center.</p><div class="heroActions"><button class="btn" onclick="showPage('riskcenter')">Open Risk Center</button><button class="btn secondary" onclick="showPage('briefbuilder')">Generate Executive Brief</button><button class="btn ghost" onclick="showPage('metrics')">Explain Metrics</button></div></div>${metrics(data.summary)}<br>${miniDeck(data.summary)}`)}${section('riskmix','Risk distribution and quality trend',`<div class="grid grid2"><div class="panel"><h3>Risk distribution</h3>${data.charts.risk}</div><div class="panel"><h3>Quality trend</h3>${data.charts.trend}</div></div>`)}${section('volume','Volume and dependencies',`<div class="grid grid2"><div class="panel"><h3>Unmatched volume</h3>${data.charts.volume}</div><div class="panel"><h3>Report dependency network</h3>${data.charts.network}</div></div>`)}${section('sankey','Sankey lineage',`<div class="panel">${data.charts.sankey}</div>`)}${section('queue','Priority queue',`<div class="panel"><h3>Priority queue</h3>${cards(data.summary.top_actions.slice(0,8))}</div>`)}`;html+=sectionShell([['overview','Overview'],['riskmix','Risk mix'],['volume','Volume'],['sankey','Sankey'],['queue','Priority queue']],body);}
if(state.page==='metrics'){html+=`<div class="hero"><h1>Metric Explanations</h1><p>Plain-English definitions for the DART metrics so leadership, analysts, and owners can interpret the same dashboard consistently.</p></div>${metricExplanations()}`;}
if(state.page==='briefing'){let body=`${section('brief-top','Briefing snapshot',`<div class="hero"><h1>My Briefing</h1><p>${esc(p.first_question||'What should I investigate first?')} This view translates reconciliation output into decision-ready priorities.</p></div>${miniDeck(data.summary)}`)}${section('focus','Recommended focus',`<div class="panel">${cards(data.summary.top_actions)}</div>`)}${section('streams','Stream summary',`<div class="panel">${simpleTable(data.summary.stream_summary)}</div>`)}`;html+=sectionShell([['brief-top','Snapshot'],['focus','Focus'],['streams','Streams']],body);}
if(state.page==='riskcenter'){let body=`${section('risk-overview','Risk overview',`<div class="hero"><h1>Risk Center</h1><p>Concentrated view of critical findings, risk drivers, impact heatmaps, and risk distribution.</p></div>${miniDeck(data.summary)}`)}${section('risk-heat','Heatmap and drivers',`<div class="grid grid2"><div class="panel"><h3>Average impact heatmap</h3>${data.charts.heatmap}</div><div class="panel"><h3>Top risk drivers</h3>${data.charts.drivers}</div></div>`)}${section('pipeline','Remediation pipeline',`<div class="grid grid2"><div class="panel"><h3>Pipeline</h3>${data.charts.pipeline}</div><div class="panel"><h3>Ownership mix</h3>${data.charts.ownership}</div></div>`)}${section('risk-list','Critical and elevated findings',`<div class="panel">${table(data.issues,true)}</div>`)}`;html+=sectionShell([['risk-overview','Overview'],['risk-heat','Heatmap'],['pipeline','Pipeline'],['risk-list','Findings']],body);}
if(state.page==='insights'){let body=`${section('insight-top','Insight overview',`<div class="hero"><h1>AI Insights</h1><p>Decision-support narratives generated from the current filtered data. These insights use deterministic local logic unless you ask DART Copilot.</p></div>${miniDeck(data.summary)}`)}${section('insight-cards','Insight cards',insightsHtml(aiInsights(data)))}${section('evidence','Evidence behind insights',`<div class="panel">${cards(data.summary.top_actions.slice(0,8))}</div>`)}`;html+=sectionShell([['insight-top','Overview'],['insight-cards','Cards'],['evidence','Evidence']],body);}
if(state.page==='scorecards'){html+=`<div class="hero"><h1>Stream Scorecards</h1><p>One executive card per stream, showing quality, volume, criticality, and impact.</p></div>${scorecards(data.summary.stream_summary)}`;}
if(state.page==='simulator'){let body=`${section('sim-select','Select field',`<div class="hero"><h1>Impact Simulator</h1><p>Select a field and simulate the downstream risk narrative for reports, KPIs, and business owners.</p></div><div class="panel"><div class="field"><label>Field</label><select id="simField" onchange='renderSimulation(${JSON.stringify(data.rows).replace(/'/g,"'")})'>${data.rows.map((r,i)=>`<option value="${i}">${esc(r.Stream)} · ${esc(r['NCH Target Column'])}</option>`).join('')}</select></div><br><button class="btn" onclick='renderSimulation(${JSON.stringify(data.rows).replace(/'/g,"'")})'>Run simulation</button></div>`)}${section('sim-result','Simulation result',`<div class="panel"><div id="simOut" class="briefText">Choose a field and run the simulator.</div></div>`)}`;html+=sectionShell([['sim-select','Select'],['sim-result','Result']],body);}
if(state.page==='actioncenter'){let body=`${section('action-overview','Action status',`<div class="hero"><h1>Remediation Center</h1><p>Seeded workflow issues are loaded automatically, so the board is useful before you add anything manually.</p></div><div class="grid grid4"><div class="panel metric"><div class="label">Open</div><div class="value">${(meta.issues||[]).filter(x=>x.Status==='Open').length}</div></div><div class="panel metric"><div class="label">In progress</div><div class="value">${(meta.issues||[]).filter(x=>x.Status==='In progress').length}</div></div><div class="panel metric"><div class="label">Blocked</div><div class="value">${(meta.issues||[]).filter(x=>x.Status==='Blocked').length}</div></div><div class="panel metric"><div class="label">Resolved</div><div class="value">${(meta.issues||[]).filter(x=>x.Status==='Resolved').length}</div></div></div>`)}${section('action-candidates','Action candidates',`<div class="panel">${table(data.issues,true)}</div>`)}${section('board','Workflow board',`<div class="panel">${kanban(meta.issues)}</div>`)}`;html+=sectionShell([['action-overview','Status'],['action-candidates','Candidates'],['board','Board']],body);}
if(state.page==='impactexplorer'){let body=`${section('impact-overview','Impact overview',`<div class="hero"><h1>Impact Explorer</h1><p>Explore how field-level reconciliation findings connect to reports, KPIs, owners, and decision needs.</p></div>${miniDeck(data.summary)}`)}${section('impact-chart','Mapped impact chart',`<div class="panel">${data.charts.impact}</div>`)}${section('impact-map','Impact mapping',impactEditor(meta.impacts,data.rows))}`;html+=sectionShell([['impact-overview','Overview'],['impact-chart','Chart'],['impact-map','Mapping']],body);}
if(state.page==='catalog'){let body=`${section('catalog-edit','Editable catalog',`<div class="hero"><h1>Mapping Catalog</h1><p>Add, edit, validate, and delete field mappings. Seeded mappings are created from the highest-impact fields.</p></div>${lineageEditor(meta.connections)}`)}`;html+=sectionShell([['catalog-edit','Catalog']],body);}
if(state.page==='lineage'){let body=`${section('lineage-network','Network',`<div class="hero"><h1>Data Lineage</h1><p>Make source-to-target relationships explicit so quality findings can be reused across reports, KPIs, and governance steps.</p></div><div class="panel">${data.charts.network}</div>`)}${section('lineage-sankey','Sankey visualization',`<div class="panel">${data.charts.sankey}</div>`)}`;html+=sectionShell([['lineage-network','Network'],['lineage-sankey','Sankey']],body);}
if(state.page==='governance'){let body=`${section('gov-overview','Governance overview',`<div class="hero"><h1>Governance Center</h1><p>Track controls, owners, evidence, cadence, and decision readiness for DART findings.</p></div><div class="grid grid4"><div class="panel metric"><div class="label">Controls</div><div class="value">${(meta.governance||[]).length}</div><div class="sub">Saved governance items</div></div><div class="panel metric"><div class="label">Active</div><div class="value">${(meta.governance||[]).filter(x=>x.Status==='Active').length}</div><div class="sub">Controls in motion</div></div><div class="panel metric"><div class="label">Mappings</div><div class="value">${(meta.connections||[]).length}</div><div class="sub">Lineage catalog rows</div></div><div class="panel metric"><div class="label">Issues</div><div class="value">${(meta.issues||[]).length}</div><div class="sub">Workflow items</div></div></div>`)}${section('gov-controls','Control register',governanceEditor(meta.governance))}${section('gov-links','Governance evidence links',`<div class="grid grid3"><div class="panel"><h3>Lineage evidence</h3><p class="muted">Use Mapping Catalog statuses as evidence that source-to-target lineage has been reviewed.</p><button class="btn secondary" onclick="showPage('catalog')">Open Mapping Catalog</button></div><div class="panel"><h3>Impact evidence</h3><p class="muted">Use Impact Explorer mappings to document downstream report/KPI exposure.</p><button class="btn secondary" onclick="showPage('impactexplorer')">Open Impact Explorer</button></div><div class="panel"><h3>Issue evidence</h3><p class="muted">Use Remediation Center workflow items to show owner assignment and remediation progress.</p><button class="btn secondary" onclick="showPage('actioncenter')">Open Remediation Center</button></div></div>`)}`;html+=sectionShell([['gov-overview','Overview'],['gov-controls','Controls'],['gov-links','Evidence links']],body);}
if(state.page==='explorer'){html+=`<div class="hero"><h1>System Explorer</h1><p>Search and inspect System A to System B relationships with risk tier and volume context next to each mapped field.</p></div><div class="panel"><h3>${intFmt(data.summary.rows)} fields in view</h3>${table(data.rows)}</div>`;}
if(state.page==='quality'){let body=`${section('quality-signal','Quality signal',`<div class="hero"><h1>Quality Analytics</h1><p>Compare equal-weight field averages, volume-weighted match rates, and distribution-level patterns.</p></div>${miniDeck(data.summary)}`)}${section('distribution','Distribution and heatmap',`<div class="grid grid2"><div class="panel"><h3>Distribution</h3>${data.charts.hist}</div><div class="panel"><h3>Impact heatmap</h3>${data.charts.heatmap}</div></div>`)}${section('weighted','Weighted performance',`<div class="panel"><h3>Reconciliation volume</h3>${data.charts.weighted}</div>`)}`;html+=sectionShell([['quality-signal','Signal'],['distribution','Distribution'],['weighted','Reconciliation volume']],body);}
if(state.page==='compare'){let body=`${section('compare-charts','Comparison charts',`<div class="hero"><h1>Stream Comparison</h1><p>Compare streams by average quality, weighted quality, unmatched volume, and risk concentration.</p></div><div class="grid grid2"><div class="panel"><h3>Field average vs weighted match</h3>${data.charts.compare}</div><div class="panel"><h3>Classification mix</h3>${data.charts.classmix}</div></div>`)}${section('compare-table','Comparison table',`<div class="panel">${simpleTable(data.compare_table)}</div>`)}`;html+=sectionShell([['compare-charts','Charts'],['compare-table','Table']],body);}
if(state.page==='briefbuilder'){let body=`${section('brief-builder','Generated brief',`<div class="hero"><h1>Executive Brief Builder</h1><p>Create a leadership-ready brief from the current filtered scope.</p><div class="heroActions"><button class="btn" onclick="copyBrief()">Copy brief</button><button class="btn secondary" onclick="showPage('settings')">Update briefing settings</button></div></div>${miniDeck(data.summary)}<br><div class="panel"><h3>Generated brief</h3><div id="briefText" class="briefText">${esc(executiveBrief(data,meta))}</div></div>`)}`;html+=sectionShell([['brief-builder','Brief']],body);}
if(state.page==='copilot'){let body=`${section('chat-area','Conversation',`<div class="hero"><h1>DART Copilot</h1><p>Ask DART Copilot questions about the filtered dataset. Without an API key, it still returns deterministic local analysis.</p></div><div class="panel"><h3>Conversation</h3><div class="chat">${meta.chat.map(m=>`<div class="bubble ${m.role==='user'?'user':'assistant'}">${esc(m.content)}</div>`).join('')}</div><br><div class="filterGrid" style="grid-template-columns:1fr auto"><input id="prompt" placeholder="What are the top risks and why?"><button class="btn" onclick="askCopilot()">Send</button></div></div>`)}${section('chat-context','Current context',`<div class="panel">${miniDeck(data.summary)}<br><h3>Suggested prompts</h3><button class="btn secondary" onclick="quickPrompt('Summarize the biggest quality risks in the current scope')">Summarize risks</button> <button class="btn secondary" onclick="quickPrompt('What should I investigate first?')">Next investigation</button> <button class="btn secondary" onclick="quickPrompt('Explain health score, field average, and event weighted match rate')">Explain metrics</button></div>`)}`;html+=sectionShell([['chat-area','Chat'],['chat-context','Context']],body);}
if(state.page==='data'){let body=`${section('data-upload','Upload dataset',`<div class="hero"><h1>Data Management</h1><p>The bundled All Streams workbook is loaded automatically and enriched with the CMS NCH-to-STTM mapping workbook. You can still upload a replacement reconciliation file when needed.</p></div><div class="grid grid2"><div class="panel"><h3>Upload dataset</h3><input type="file" id="file"><br><br><button class="btn" onclick="uploadFile()">Use uploaded data</button> <button class="btn secondary" onclick="restoreDemo()">Reload Excel files</button><br><br><a class="btn ghost" href="/download/current.csv">Download current data</a><p class="muted">Source: ${esc(meta.source)} · Rows: ${intFmt(meta.rows)} · Columns: ${intFmt(meta.columns)}</p></div><div class="panel"><h3>Detected structure</h3>${simpleTable(meta.profile)}</div></div>`)}${section('data-preview','Preview',`<div class="panel"><h3>Preview</h3>${table(data.rows)}</div>`)}`;html+=sectionShell([['data-upload','Upload'],['data-preview','Preview']],body);}
if(state.page==='profile'){html+=profilePage(meta);}if(state.page==='settings'){html+=settingsPage(meta);}const appEl=document.getElementById('app');appEl.classList.remove('page-enter');appEl.innerHTML=html;runInjectedScripts(appEl);requestAnimationFrame(()=>appEl.classList.add('page-enter'));}
function profilePage(meta){const u=meta.auth_user||{};return `<div class="hero"><h1>Profile</h1><p>Manage your local DART prototype profile. Credentials are stored in <code>${esc(u.storage_file||'dart_users.json')}</code> for easy local testing only.</p><div class="heroActions"><button class="btn secondary" onclick="logoutUser()">Logout</button></div></div><form class="panel" id="profileForm"><div class="grid grid2"><div class="field"><label>Username</label><input value="${esc(u.username||'')}" disabled></div><div class="field"><label>Display name</label><input id="profileName" value="${esc(u.display_name||'')}"></div><div class="field"><label>Email</label><input id="profileEmail" value="${esc(u.email||'')}"></div><div class="field"><label>Organization</label><input id="profileOrg" value="${esc(u.organization||'')}"></div><div class="field"><label>Role</label><input id="profileRole" value="${esc(u.role||'')}"></div><div class="field"><label>New password</label><input id="profilePass" type="password" placeholder="Leave blank to keep current password"></div></div><br><button class="btn" type="submit">Save profile</button></form><br><div class="panel"><h3>Prototype storage note</h3><p class="muted">This prototype profile system stores credentials locally in plain text. It is useful for local demos, but should be replaced with proper authentication before use with sensitive data.</p></div>`;}
function settingsPage(meta){const p=meta.persona||{};const focus=(p.focus||[]).join(', ');return `<div class="hero"><h1>Workspace Settings</h1><p>Update the briefing lens and personalization settings without resetting your data.</p></div><form class="panel" id="settingsForm"><div class="grid grid2"><div class="field"><label>Program</label><select id="setProgram">${['Medicare','Medicaid','Both','Other / General'].map(x=>`<option ${p.program===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Primary role</label><select id="setRole">${['Executive / Leadership','Data / Analytics','Program / Policy','Operations','Quality / Compliance','IT / Engineering','Research','Other'].map(x=>`<option ${p.role===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Focus areas</label><input id="setFocus" value="${esc(focus)}"></div><div class="field"><label>Audience</label><select id="setAudience">${['Leadership','Myself','Analysts','Program teams','Technical teams','External stakeholders'].map(x=>`<option ${p.audience===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Detail level</label><select id="setDepth">${['Executive','Balanced','Technical'].map(x=>`<option ${p.depth===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Briefing question</label><input id="setQuestion" value="${esc(p.first_question||'')}"></div></div><br><button class="btn" type="submit">Save settings</button></form>`;}
function renderSimulation(rows){if(!rows.length){simOut.innerText='No rows available.';return;}const r=rows[Number(simField.value||0)];simOut.innerText=`Impact Simulation\n\nField: ${r.Stream} · ${r['NCH Target Column']}\nSource: ${r['NCH Target Table']}\nTarget: ${r['SS Table']}\n\nQuality Signal\n- Match rate: ${pctNum(r.MatchRate)}\n- Not matched: ${intFmt(r.NotMatchedClaims)}\n- Risk tier: ${r.RiskTier}\n- Impact score: ${r.ImpactScore}\n\nPotential Downstream Impact\n- Report: ${r.Report || r.Stream + ' Quality Report'}\n- KPI: ${r.KPI || r.Stream + ' Match KPI'}\n- Business severity: ${r.RiskTier}\n\nSuggested action\nValidate mapping logic, confirm lineage, and assign an owner if this supports a leadership KPI.`;}
async function addImpact(rows){if(!rows.length){toast('No rows to map');return;}const r=rows[Number(impactField.value)];await postJson('/api/impacts',{Stream:r.Stream,Field:r['NCH Target Column'],Report:impactReport.value,KPI:impactKpi.value,Owner:impactOwner.value,Impact:impactLevel.value,DecisionNeed:impactDecision.value});toast('Impact mapped');await render();}
async function askCopilot(){const v=prompt.value;if(!v.trim())return;await postJson('/api/copilot',{prompt:v,filters:state.filters});await render();}
async function quickPrompt(v){await postJson('/api/copilot',{prompt:v,filters:state.filters});await render();}
async function uploadFile(){const f=file.files[0];if(!f){toast('Choose a file first');return;}const fd=new FormData();fd.append('file',f);const r=await fetch('/api/upload',{method:'POST',body:fd});if(!r.ok){toast(await r.text());return;}toast('Data loaded');await render();}
async function restoreDemo(){await postJson('/api/restore-demo',{});toast('Excel files reloaded');await render();}
function copyBrief(){const txt=document.getElementById('briefText')?.innerText||'';navigator.clipboard.writeText(txt);toast('Brief copied');}
document.addEventListener('submit',e=>{if(e.target&&e.target.id==='settingsForm')saveSettings(e);if(e.target&&e.target.id==='profileForm')saveProfile(e)});
boot();
</script>
</body>
</html>
'''

# -----------------------------------------------------------------------------
# API routes
# -----------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index() -> str:
    ensure_data()
    return HTML


@app.post("/api/signup")
async def signup(request: Request) -> Any:
    payload = await request.json()
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", "")).strip()
    if not username or not password:
        return JSONResponse({"error": "Username and password are required."}, status_code=400)
    data = load_users()
    users = data.setdefault("users", {})
    if username in users:
        return JSONResponse({"error": "Username already exists."}, status_code=400)
    users[username] = {
        "password": password,
        "created_at": "local prototype",
        "profile": {
            "display_name": str(payload.get("display_name", "")).strip() or username,
            "email": str(payload.get("email", "")).strip(),
            "organization": str(payload.get("organization", "")).strip(),
            "role": str(payload.get("role", "")).strip(),
        },
    }
    save_users(data)
    STATE["current_user"] = username
    return {"status": "ok"}


@app.post("/api/login")
async def login(request: Request) -> Any:
    payload = await request.json()
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", "")).strip()
    user = load_users().get("users", {}).get(username)
    if not user or user.get("password") != password:
        return JSONResponse({"error": "Invalid username or password."}, status_code=401)
    STATE["current_user"] = username
    return {"status": "ok"}


@app.post("/api/logout")
async def logout() -> Dict[str, str]:
    STATE["current_user"] = None
    return {"status": "ok"}


@app.post("/api/profile")
async def update_profile(request: Request) -> Any:
    username = STATE.get("current_user")
    if not username:
        return JSONResponse({"error": "Not logged in."}, status_code=401)
    payload = await request.json()
    data = load_users()
    user = data.setdefault("users", {}).setdefault(username, {"password": "", "profile": {}})
    profile = user.setdefault("profile", {})
    for key in ["display_name", "email", "organization", "role"]:
        if key in payload:
            profile[key] = str(payload.get(key, "")).strip()
    new_password = str(payload.get("password", "")).strip()
    if new_password:
        user["password"] = new_password
    save_users(data)
    return {"status": "ok"}


@app.get("/api/meta")
def meta() -> Dict[str, Any]:
    ensure_data()
    df = standardize(STATE["df"])
    profile = pd.DataFrame({
        "Column": df.columns,
        "Type": [str(df[c].dtype) for c in df.columns],
        "Non-null": [int(df[c].notna().sum()) for c in df.columns],
        "Unique": [int(df[c].nunique(dropna=True)) for c in df.columns],
    }).to_dict("records")
    return {
        "auth_user": public_user(STATE.get("current_user")),
        "persona": STATE["persona"],
        "source": STATE["source"],
        "version": APP_VERSION,
        "rows": len(df),
        "columns": len(df.columns),
        "columns_list": [str(c) for c in df.columns],
        "streams": sorted(df["Stream"].astype(str).unique().tolist()) if "Stream" in df else [],
        "classes": sorted(df["Classification"].astype(str).unique().tolist()) if "Classification" in df else [],
        "tiers": sorted(df["RiskTier"].astype(str).unique().tolist()) if "RiskTier" in df else [],
        "reasons": sorted(df["Sub-Classification"].astype(str).unique().tolist()) if "Sub-Classification" in df else [],
        "profile": profile,
        "issues": STATE["issues"],
        "connections": STATE["connections"],
        "impacts": STATE["impacts"],
        "governance": STATE["governance"],
        "rules": STATE["rules"],
        "chat": STATE["chat"],
    }


@app.post("/api/persona")
async def persona(request: Request) -> Dict[str, str]:
    STATE["persona"] = await request.json()
    if not STATE["chat"]:
        STATE["chat"] = [{"role": "assistant", "content": "Hi. I am Dartboard, your DART assistant. Ask what changed, where the biggest risks are, or what to investigate next."}]
    return {"status": "ok"}


@app.post("/api/view")
async def view(request: Request) -> Dict[str, Any]:
    filters = await request.json()
    fdf = apply_filters(filters)
    charts, compare_table = build_charts(fdf)
    issues = fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]) if not fdf.empty else fdf
    return {"summary": summary_payload(fdf), "rows": clean_records(fdf, 225), "issues": clean_records(issues, 125), "charts": charts, "compare_table": compare_table}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)) -> Any:
    return JSONResponse(
        {"error": "Uploads are disabled. DART uses only All Streams June 2026.xlsx and CMS Version NCH to NCH STTM 12_13_2024.xlsx."},
        status_code=403,
    )

@app.post("/api/restore-demo")
async def reload_excel_files() -> Dict[str, str]:
    STATE["df"] = pd.DataFrame()
    STATE["source"] = ""
    STATE["snapshots"] = []
    STATE["connections"] = []
    STATE["issues"] = []
    STATE["impacts"] = []
    ensure_data()
    return {"status": "ok"}


@app.post("/api/seed-lineage")
async def seed_lineage() -> Dict[str, str]:
    STATE["connections"] = []
    seed_lineage_from_df(standardize(STATE["df"]))
    return {"status": "ok"}


@app.post("/api/lineage")
async def upsert_lineage(request: Request) -> Dict[str, str]:
    payload = await request.json()
    item = {k: str(payload.get(k, "")).strip() for k in ["id", "Stream", "SourceTable", "SourceField", "TargetTable", "TargetField", "Report", "KPI", "Owner", "Status", "RiskTier", "Notes"]}
    if not item["id"]:
        item["id"] = f"L{len(STATE['connections']) + 1}"
    found = False
    for i, existing in enumerate(STATE["connections"]):
        if str(existing.get("id")) == item["id"]:
            STATE["connections"][i] = item
            found = True
            break
    if not found:
        STATE["connections"].append(item)
    return {"status": "ok"}


@app.delete("/api/lineage/{item_id}")
async def delete_lineage(item_id: str) -> Dict[str, str]:
    STATE["connections"] = [x for x in STATE["connections"] if str(x.get("id")) != item_id]
    return {"status": "ok"}


@app.post("/api/issues")
async def add_issue(request: Request) -> Dict[str, str]:
    payload = await request.json()
    payload["id"] = payload.get("id") or f"I{len(STATE['issues']) + 1}"
    STATE["issues"].append(payload)
    return {"status": "ok"}


@app.post("/api/impacts")
async def add_impact(request: Request) -> Dict[str, str]:
    payload = await request.json()
    payload["id"] = payload.get("id") or f"M{len(STATE['impacts']) + 1}"
    STATE["impacts"].append(payload)
    return {"status": "ok"}


@app.post("/api/governance")
async def upsert_governance(request: Request) -> Dict[str, str]:
    payload = await request.json()
    item = {k: str(payload.get(k, "")).strip() for k in ["id", "Control", "Owner", "Cadence", "Status", "Evidence"]}
    if not item["id"]:
        item["id"] = f"G{len(STATE['governance']) + 1}"
    found = False
    for i, existing in enumerate(STATE["governance"]):
        if str(existing.get("id")) == item["id"]:
            STATE["governance"][i] = item
            found = True
            break
    if not found:
        STATE["governance"].append(item)
    return {"status": "ok"}


@app.delete("/api/governance/{item_id}")
async def delete_governance(item_id: str) -> Dict[str, str]:
    STATE["governance"] = [x for x in STATE["governance"] if str(x.get("id")) != item_id]
    return {"status": "ok"}


@app.post("/api/rules")
async def add_rule(request: Request) -> Dict[str, str]:
    STATE["rules"].append(await request.json())
    return {"status": "ok"}


@app.post("/api/copilot")
async def copilot(request: Request) -> Dict[str, str]:
    payload = await request.json()
    prompt = str(payload.get("prompt", "")).strip()
    filters = payload.get("filters", {})
    fdf = apply_filters(filters)
    if prompt:
        STATE["chat"].append({"role": "user", "content": prompt})
    context = {
        "persona": STATE["persona"],
        "rows": len(fdf),
        "field_average": field_average(fdf) if not fdf.empty else None,
        "weighted_rate": weighted_rate(fdf) if not fdf.empty else None,
        "top_actions": clean_records(fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]), 15) if not fdf.empty else [],
        "stream_summary": summary_payload(fdf).get("stream_summary", []),
        "lineage_count": len(STATE.get("connections", [])),
        "impact_count": len(STATE.get("impacts", [])),
        "governance_count": len(STATE.get("governance", [])),
    }
    api_key = os.getenv("GROQ_API_KEY", "")
    model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    system = "You are Dartboard, the DART healthcare data intelligence assistant. Be concise, executive-ready, and data-driven. Use only supplied context. Never invent statistics. This is decision support, not clinical advice."
    if OpenAI and api_key:
        try:
            client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
            messages = [{"role": "system", "content": system + "\n\nCURRENT CONTEXT:\n" + json.dumps(context, default=str)}, *STATE["chat"][-8:]]
            res = client.chat.completions.create(model=model, temperature=0.2, messages=messages)
            answer = res.choices[0].message.content or "I could not generate an answer from the current context."
        except Exception as exc:
            answer = f"AI service error: {exc}\n\nLocal result: field average is {pct(field_average(fdf))}; overall average match is {pct(weighted_rate(fdf))}."
    else:
        if fdf.empty:
            answer = "No rows match the current filters. Adjust scope or upload a dataset."
        else:
            top = fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(3)
            bullets = "\n".join([f"- {r['Stream']} · {r['NCH Target Column']}: {pct(r['MatchRate'])} match, {fmt_int(r['NotMatchedClaims'])} unmatched, {r['RiskTier']} risk" for _, r in top.iterrows()])
            answer = f"Current scope has {len(fdf):,} fields. Field average is {pct(field_average(fdf))}; overall average match is {pct(weighted_rate(fdf))}. The workspace also has {len(STATE.get('connections', []))} lineage mappings, {len(STATE.get('impacts', []))} impact mappings, and {len(STATE.get('governance', []))} governance controls. Top action candidates:\n{bullets}\n\nAdd GROQ_API_KEY for model-backed Dartboard analysis."
    STATE["chat"].append({"role": "assistant", "content": answer})
    return {"status": "ok"}


@app.get("/download/current.csv")
def download_current() -> StreamingResponse:
    ensure_data()
    csv = standardize(STATE["df"]).to_csv(index=False).encode("utf-8")
    return StreamingResponse(io.BytesIO(csv), media_type="text/csv", headers={"Content-Disposition": "attachment; filename=dart_current.csv"})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=int(os.getenv("PORT", "8000")))
