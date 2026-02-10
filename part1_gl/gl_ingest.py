from __future__ import annotations

import os
from typing import Dict, Tuple

import pandas as pd

from common.sqlite_utils import clear_table, connect, executemany, table_exists


GL_COLUMNS = [
    "Month",
    "Date",
    "GL Account",
    "Type",
    "Property Name",
    "Property Street Address 1",
    "Property Street Address 2",
    "Debit",
    "Credit",
    "Balance",
]


def _coerce_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _to_iso_date(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, format="%m/%d/%Y", errors="coerce")
    return dt.dt.strftime("%Y-%m-%d")


def _to_month_start_iso(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series.astype(str).str.strip(), format="%b %Y", errors="coerce")
    dt = dt.dt.to_period("M").dt.to_timestamp()
    return dt.dt.strftime("%Y-%m-%d")


def ensure_schema(conn, gl_raw_table: str) -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {gl_raw_table} (
            month_start TEXT,
            txn_date TEXT,
            gl_account TEXT,
            gl_type TEXT,
            property_name TEXT,
            property_street_address_1 TEXT,
            property_street_address_2 TEXT,
            debit REAL,
            credit REAL,
            balance REAL,
            investor TEXT,
            owner TEXT,
            acquired TEXT,
            investor_type TEXT,
            categorization TEXT,
            gl_mapping_type TEXT,
            cash_categorization TEXT,
            cash_type_mapping TEXT
        );
        """
    )
    conn.commit()

def reset_raw_table(conn, gl_raw_table: str) -> None:
    conn.execute(f"DROP TABLE IF EXISTS {gl_raw_table};")
    ensure_schema(conn, gl_raw_table)
    conn.commit()

def ingest_gl_csv_to_raw(
    db_path: str,
    gl_raw_table: str,
    csv_path: str,
    rows_to_skip_after_header: int,
) -> None:
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"GL CSV not found: {csv_path}")

    skiprows = list(range(1, 1 + int(rows_to_skip_after_header)))

    df = pd.read_csv(
        csv_path,
        header=0,
        skiprows=skiprows,
        dtype=str,
        keep_default_na=False,
    )

    missing = [c for c in GL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"GL CSV missing required columns: {missing}")

    df = df[GL_COLUMNS].copy()
    df["Month"] = df["Month"].astype(str).str.strip()
    df["Date"] = df["Date"].astype(str).str.strip()


    df["debit"] = _coerce_numeric(df["Debit"])
    df["credit"] = _coerce_numeric(df["Credit"])
    df["balance"] = _coerce_numeric(df["Balance"])

    df["month_start"] = _to_month_start_iso(df["Month"])
    df["txn_date"] = _to_iso_date(df["Date"])

    df["gl_account"] = df["GL Account"].astype(str).str.strip()
    df["gl_type"] = df["Type"].astype(str).str.strip()
    df["property_name"] = df["Property Name"].astype(str).str.strip()
    df["property_street_address_1"] = df["Property Street Address 1"].astype(str).str.strip()
    df["property_street_address_2"] = df["Property Street Address 2"].astype(str).str.strip()

    df["investor"] = None
    df["owner"] = None
    df["acquired"] = None
    df["investor_type"] = None
    df["categorization"] = None
    df["gl_mapping_type"] = None
    df["cash_categorization"] = None
    df["cash_type_mapping"] = None

    insert_cols = [
        "month_start",
        "txn_date",
        "gl_account",
        "gl_type",
        "property_name",
        "property_street_address_1",
        "property_street_address_2",
        "debit",
        "credit",
        "balance",
        "investor",
        "owner",
        "acquired",
        "investor_type",
        "categorization",
        "gl_mapping_type",
        "cash_categorization",
        "cash_type_mapping",
    ]

    rows = [tuple(x) for x in df[insert_cols].itertuples(index=False, name=None)]

    conn = connect(db_path)
    try:
        reset_raw_table(conn, gl_raw_table)

        sql = f"""
        INSERT INTO {gl_raw_table} (
            month_start, txn_date, gl_account, gl_type, property_name,
            property_street_address_1, property_street_address_2,
            debit, credit, balance,
            investor, owner, acquired, investor_type, categorization, gl_mapping_type, cash_categorization, cash_type_mapping
        ) VALUES (
            ?, ?, ?, ?, ?,
            ?, ?,
            ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?
        );
        """

        executemany(conn, sql, rows)
        conn.commit()
    finally:
        conn.close()
