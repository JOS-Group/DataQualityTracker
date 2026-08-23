import re
from pathlib import Path

import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"

_CACHE: dict[str, tuple[float, pd.DataFrame]] = {}


def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def find_col(df: pd.DataFrame, choices: list[str]):
    exact = {str(c).lower().strip(): c for c in df.columns}
    for x in choices:
        if x.lower().strip() in exact:
            return exact[x.lower().strip()]
    for c in df.columns:
        for x in choices:
            if x.lower() in str(c).lower():
                return c
    return None


def standardize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = norm_cols(df)

    aliases = {
        "Stream": ["Stream", "Claim Stream", "Domain"],
        "NCH Target Table": ["NCH Target Table", "NCH Table", "System A Table", "Source Table"],
        "NCH Target Column": ["NCH Target Column", "NCH Column", "System A Column", "Source Column"],
        "SS Table": ["SS Table", "Shared Systems Table", "System B Table", "Target Table"],
        "SS Column": ["SS Column", "Shared Systems Column", "System B Column", "Target Column"],
        "MatchedClaims": ["MatchedClaims", "Matched Claims", "Matched Events", "Matched"],
        "NotMatchedClaims": ["NotMatchedClaims", "Not Matched Claims", "Not-Matched Claims", "Unmatched"],
        "MatchRate": ["MatchRate", "Match Rate", "Match %", "Match Percentage"],
        "Classification": ["Classification", "ClassificationClean", "Class"],
        "Disposition": ["Disposition", "DispositionClean", "Action"],
        "Recommendation": ["Recommendation", "Recommended Action"],
        "Comments": ["Comments", "Comment", "Notes"],
    }
    ren = {}
    for target, opts in aliases.items():
        c = find_col(df, opts)
        if c and c != target:
            ren[c] = target
    df = df.rename(columns=ren)

    for c in ["MatchedClaims", "NotMatchedClaims", "MatchRate"]:
        if c in df:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )

    if "MatchRate" in df and df["MatchRate"].dropna().size and df["MatchRate"].dropna().max() > 1.5:
        df["MatchRate"] /= 100

    for c in ["MatchedClaims", "NotMatchedClaims"]:
        if c not in df:
            df[c] = np.nan

    if "MatchRate" not in df:
        total = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
        df["MatchRate"] = np.where(total > 0, df["MatchedClaims"] / total, np.nan)

    if "Classification" not in df:
        df["Classification"] = np.where(
            df["MatchRate"] >= 0.9999, "Match",
            np.where(df["MatchRate"].notna(), "Different", "Missing"),
        )

    for c in ["Stream", "NCH Target Table", "NCH Target Column", "SS Table", "SS Column"]:
        if c not in df:
            df[c] = ""

    df["Comparable"] = df["Classification"].astype(str).str.strip().str.lower().isin(["match", "different"])

    if "Disposition" in df:
        disp = df["Disposition"].astype(str).str.lower()
        df["NeedsChange"] = disp.str.contains("action|change")
    else:
        df["NeedsChange"] = False
    df["NeedsChange"] |= ~df["Comparable"]

    df["TotalClaims"] = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
    return df


def field_average(df: pd.DataFrame):
    x = df[df["Comparable"] & df["MatchRate"].notna()]
    return x["MatchRate"].mean() if len(x) else np.nan


def weighted_rate(df: pd.DataFrame):
    m = df["MatchedClaims"].fillna(0).sum()
    u = df["NotMatchedClaims"].fillna(0).sum()
    return m / (m + u) if m + u else np.nan


def fmt_pct(x):
    return "—" if x is None or pd.isna(x) else f"{x:.1%}"


def fmt_num(x):
    return "—" if x is None or pd.isna(x) else f"{x:,.0f}"


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    return df.replace({np.nan: None})


def records(df: pd.DataFrame) -> list[dict]:
    return _clean(df).to_dict("records")


def records_with_id(df: pd.DataFrame) -> list[dict]:
    out = []
    for idx, row in _clean(df).iterrows():
        d = row.to_dict()
        d["_id"] = idx
        out.append(d)
    return out


def demo_data() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    meta = [("PROF", "MCS"), ("INP/SNF", "FISS"), ("OUT", "FISS"),
            ("HH", "FISS"), ("HSPC", "FISS"), ("DME", "VMS")]
    rows = []
    for stream, ss in meta:
        for i in range(20):
            rate = float(rng.uniform(.38, .99))
            matched = int(rng.integers(3000, 200000))
            total = max(int(matched / max(rate, .05)), matched)
            rows.append({
                "Stream": stream,
                "NCH Target Table": f"{stream}_TABLE_{i % 5 + 1}",
                "NCH Target Column": f"{stream}_FIELD_{i + 1}",
                "SS Table": f"{ss}_TABLE_{i % 5 + 1}",
                "SS Column": f"{ss}_FIELD_{i + 1}",
                "MatchedClaims": matched,
                "NotMatchedClaims": total - matched,
                "MatchRate": matched / total,
                "Classification": "Match" if rate >= .9 else ("Different" if rate >= .55 else "Missing"),
                "Disposition": "Good" if rate >= .9 else "Action needed",
                "Recommendation": "Review mapping" if rate < .9 else "No action",
                "Comments": "Interactive demo record",
            })
    return standardize(pd.DataFrame(rows))


def load_df_from_path(path: str) -> pd.DataFrame:
    p = Path(path)
    mtime = p.stat().st_mtime
    cached = _CACHE.get(str(p))
    if cached and cached[0] == mtime:
        return cached[1]
    raw = pd.read_csv(p) if p.suffix.lower() == ".csv" else pd.read_excel(p)
    df = standardize(raw)
    _CACHE[str(p)] = (mtime, df)
    return df


def evaluate_rule(df: pd.DataFrame, rule) -> pd.DataFrame:
    col, cond, val = rule.field, rule.condition, rule.value
    if col not in df.columns:
        return df.iloc[0:0]
    try:
        if cond == "is empty / null":
            return df[df[col].isna() | df[col].astype(str).str.strip().eq("")]
        if cond == "is not empty":
            return df[df[col].notna() & df[col].astype(str).str.strip().ne("")]
        if cond == "equals value":
            return df[df[col].astype(str).str.lower() == str(val).lower()]
        if cond == "contains text":
            return df[df[col].astype(str).str.contains(val, case=False, na=False)]
        if cond == "above number":
            return df[pd.to_numeric(df[col], errors="coerce") > float(val)]
        if cond == "below number":
            return df[pd.to_numeric(df[col], errors="coerce") < float(val)]
        if cond == "match rate below threshold":
            return df[df["MatchRate"] < float(val) / 100]
        if cond == "needs action":
            return df[df["NeedsChange"]]
    except Exception:
        pass
    return df.iloc[0:0]
