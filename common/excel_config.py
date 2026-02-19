from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd


@dataclass(frozen=True)
class GeneralConfig:
    gl_location: str
    gl_file_name: str
    output_location: str
    statement_thru_date: str
    studio_market: float
    one_bed_market: float
    two_bed_market: float
    three_bed_market: float

@dataclass(frozen=True)
class SetupConfig:
    general: GeneralConfig
    investors: List[str]
    run_config: pd.DataFrame
    investor_table: pd.DataFrame
    gl_mapping: pd.DataFrame

def _normalize_key(value: str) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(":"):
        s = s[:-1].strip()
    return s.lower()

def _read_general_config(xlsx_path: str) -> GeneralConfig:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="General Config",
        header=None,
        engine="openpyxl",
    )
    df = df.dropna(how="all")
    if df.shape[1] < 2:
        raise ValueError("General Config sheet must have at least two columns.")

    kv: Dict[str, str] = {}
    for _, row in df.iterrows():
        k = _normalize_key(row.iloc[0])
        v = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
        if k:
            kv[k] = v

    def get_required(key: str) -> str:
        nk = _normalize_key(key)
        if nk not in kv or not kv[nk]:
            raise ValueError(f"Missing required General Config value for: {key}")
        return kv[nk]

    return GeneralConfig(
        gl_location=get_required("GL Location"),
        gl_file_name=get_required("GL File Name"),
        output_location=get_required("Output Location"),
        statement_thru_date=get_required("Statement Thru Date"),
        studio_market=float(get_required("Studio Market")),
        one_bed_market=float(get_required("1-Bed Market")),
        two_bed_market=float(get_required("2-Bed Market")),
        three_bed_market=float(get_required("3-Bed Market")),
    )


def _read_run_config(xlsx_path: str) -> Tuple[List[str], pd.DataFrame]:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="Run Config",
        engine="openpyxl",
    )

    required = ["Investor", "Owner", "% Ownership"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Run Config missing required columns: {missing}")

    out = df.copy()
    out["Investor"] = out["Investor"].astype(str).str.strip()
    out["Owner"] = out["Owner"].astype(str).str.strip()

    def _normalize_pct(v) -> float:
        if pd.isna(v):
            raise ValueError("Run Config % Ownership contains blank values.")
        if isinstance(v, (int, float)):
            x = float(v)
            if 0 < x <= 1:
                return x * 100.0
            return x
        s = str(v).strip()
        if s.endswith("%"):
            s = s[:-1].strip()
        x = float(s)
        if 0 < x <= 1:
            return x * 100.0
        return x

    out["pct_ownership"] = out["% Ownership"].apply(_normalize_pct)

    bad = out[(out["pct_ownership"] <= 0) | (out["pct_ownership"] > 100)]
    if not bad.empty:
        raise ValueError("Run Config % Ownership must be >0 and <=100 for all rows.")

    out = out[(out["Investor"] != "") & (out["Owner"] != "")].copy()

    # Drop duplicate Investor + Owner pairs, keep first
    out = out.drop_duplicates(subset=["Investor", "Owner"], keep="first").copy()

    # Validate Owner sums equal exactly 100.0, otherwise skip that Owner
    sums = out.groupby("Owner", as_index=False)["pct_ownership"].sum()
    valid_owners = set(sums.loc[sums["pct_ownership"] == 100.0, "Owner"].tolist())
    out = out[out["Owner"].isin(valid_owners)].copy()

    investors = sorted(out["Investor"].dropna().astype(str).str.strip().unique().tolist())
    if not investors:
        raise ValueError("Run Config contains no valid Investor rows after validation.")

    return investors, out[["Investor", "Owner", "pct_ownership"]].copy()

def _read_investor_table(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="Investor Table",
        engine="openpyxl",
    )

    required = ["Investor", "% Ownership", "Owner", "Property Name", "Property", "Acquired", "Type"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Investor Table missing required columns: {missing}")

    df = df[required].copy()

    df["Investor"] = df["Investor"].astype(str).str.strip()
    df["Owner"] = df["Owner"].astype(str).str.strip()
    df["Property Name"] = df["Property Name"].astype(str).str.strip()
    df["Property"] = df["Property"].astype(str).str.strip()

    def _normalize_pct(v) -> float:
        if pd.isna(v):
            raise ValueError("Investor Table % Ownership contains blank values.")
        if isinstance(v, (int, float)):
            x = float(v)
            if 0 < x <= 1:
                return x * 100.0
            return x
        s = str(v).strip()
        if s.endswith("%"):
            s = s[:-1].strip()
        x = float(s)
        if 0 < x <= 1:
            return x * 100.0
        return x

    df["pct_ownership"] = df["% Ownership"].apply(_normalize_pct)

    bad = df[(df["pct_ownership"] <= 0) | (df["pct_ownership"] > 100)]
    if not bad.empty:
        raise ValueError("Investor Table % Ownership must be >0 and <=100 for all rows.")

    # Validation: per Owner + Property Name, investor splits must total exactly 100.0
    sums = df.groupby(["Owner", "Property Name"], as_index=False)["pct_ownership"].sum()
    bad_groups = sums[sums["pct_ownership"] != 100.0]
    if not bad_groups.empty:
        examples = bad_groups.head(25).to_dict(orient="records")
        raise ValueError(f"Investor Table % Ownership must total 100.0 per Owner + Property Name. Examples: {examples}")

    return df

def _read_gl_mapping(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="GL Mapping",
        engine="openpyxl",
    )

    required = ["GL Account", "Categorization", "GL Type", "Cash Categorization", "Cash Type"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"GL Mapping missing required columns: {missing}")

    df = df[required].copy()
    df["GL Account"] = df["GL Account"].astype(str).str.strip()
    df["Categorization"] = df["Categorization"].astype(str).str.strip()
    df["GL Type"] = df["GL Type"].astype(str).str.strip()
    df["Cash Categorization"] = df["Cash Categorization"].astype(str).str.strip()
    df["Cash Type"] = df["Cash Type"].astype(str).str.strip()

    return df


def load_setup_config(xlsx_path: str) -> SetupConfig:
    general = _read_general_config(xlsx_path)
    investors, run_config = _read_run_config(xlsx_path)
    investor_table = _read_investor_table(xlsx_path)
    gl_mapping = _read_gl_mapping(xlsx_path)

    return SetupConfig(
        general=general,
        investors=investors,
        run_config=run_config,
        investor_table=investor_table,
        gl_mapping=gl_mapping,
    )

