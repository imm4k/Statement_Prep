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


def _read_investors(xlsx_path: str) -> List[str]:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="Run Config",
        engine="openpyxl",
    )
    if "Investor" not in df.columns:
        raise ValueError("Run Config must include a column named Investor.")

    investors = []
    for v in df["Investor"].tolist():
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            investors.append(s)

    if not investors:
        raise ValueError("Run Config contains no Investor values.")

    return investors


def _read_investor_table(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(
        xlsx_path,
        sheet_name="Investor Table",
        engine="openpyxl",
    )

    required = ["Investor", "Property", "Property Name", "Owner", "Acquired", "Type"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Investor Table missing required columns: {missing}")

    df = df[required].copy()
    df["Property"] = df["Property"].astype(str).str.strip()
    df["Property Name"] = df["Property Name"].astype(str).str.strip()

    dup = df["Property Name"].duplicated(keep=False)
    if dup.any():
        dups = sorted(df.loc[dup, "Property Name"].unique().tolist())
        raise ValueError(
            "Investor Table requires unique Property Name values for the current design. "
            f"Duplicates found: {dups}"
        )

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
    investors = _read_investors(xlsx_path)
    investor_table = _read_investor_table(xlsx_path)
    gl_mapping = _read_gl_mapping(xlsx_path)

    return SetupConfig(
        general=general,
        investors=investors,
        investor_table=investor_table,
        gl_mapping=gl_mapping,
    )
