from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# Ensure project root is on sys.path so "common" imports work when running via VS Code Play
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config  # part1_gl/config.py
from common.excel_config import load_setup_config
from common.sqlite_utils import connect


TARGET_PROPERTY_NAME = "CPM Luca/Jamie/David 1 - Cosac - 1585upr6B"
TARGET_CATEGORIZATION = "Rent"


def _print_df(title: str, df: pd.DataFrame, max_rows: int = 25) -> None:
    print("\n" + "=" * 90)
    print(title)
    print("=" * 90)
    if df.empty:
        print("(no rows)")
        return
    with pd.option_context("display.max_rows", max_rows, "display.max_columns", 200, "display.width", 200):
        print(df.head(max_rows).to_string(index=False))


def main() -> None:
    print("=" * 90)
    print("debug.py: Verify mapping and aggregation for one property + categorization")
    print("=" * 90)

    print(f"PROJECT_ROOT: {PROJECT_ROOT}")
    print(f"Python: {sys.version.split()[0]}")
    print(f"excel_config.py imported from: {Path(sys.modules['common.excel_config'].__file__).resolve()}")

    setup_xlsx = Path(config.SETUP_XLSX_PATH)
    db_path = Path(config.SQLITE_PATH)

    print(f"Setup XLSX: {setup_xlsx}")
    print(f"SQLite DB:  {db_path}")

    setup = load_setup_config(str(setup_xlsx))
    print(f"Run investors count (Run Config): {len(setup.investors)}")
    print(f"Investor Table rows: {len(setup.investor_table)}")
    print(f"GL Mapping rows: {len(setup.gl_mapping)}")

    # Investor Table rows for this property name
    inv = setup.investor_table.copy()
    inv = inv[inv["Property Name"].astype(str).str.strip() == TARGET_PROPERTY_NAME].copy()

    if "pct_ownership" in inv.columns:
        inv["pct_ownership"] = inv["pct_ownership"].astype(float)
        inv_sum = inv.groupby(["Owner", "Property Name"], as_index=False)["pct_ownership"].sum()
    else:
        inv_sum = pd.DataFrame()

    _print_df("Investor Table rows for TARGET_PROPERTY_NAME", inv)
    _print_df("Investor Table ownership sum by Owner + Property Name (should equal 100.0)", inv_sum)

    # GL Mapping check for Rent categorization
    glmap = setup.gl_mapping.copy()
    glmap = glmap[glmap["Categorization"].astype(str).str.strip() == TARGET_CATEGORIZATION].copy()
    _print_df("GL Mapping rows where Categorization == 'Rent' (sanity)", glmap)

    conn = connect(str(db_path))
    try:
        # 1) Show distinct categorizations in raw for this property
        df = pd.read_sql_query(
            f"""
            SELECT
                property_name,
                categorization,
                gl_mapping_type,
                cash_categorization,
                cash_type_mapping,
                COUNT(*) AS row_count,
                SUM(debit) AS sum_debit,
                SUM(credit) AS sum_credit,
                SUM(debit) - SUM(credit) AS net
            FROM {config.GL_RAW_TABLE}
            WHERE property_name = ?
            GROUP BY
                property_name,
                categorization,
                gl_mapping_type,
                cash_categorization,
                cash_type_mapping
            ORDER BY row_count DESC;
            """,
            conn,
            params=(TARGET_PROPERTY_NAME,),
        )
        _print_df("gl_raw: rollup by categorization and mapping fields for TARGET_PROPERTY_NAME", df, max_rows=200)

        # 2) Pull the raw rows for Rent for this property (so we can see month_start/txn_date)
        df_raw_rent = pd.read_sql_query(
            f"""
            SELECT
                month_start,
                txn_date,
                gl_account,
                gl_type,
                debit,
                credit,
                (debit - credit) AS net,
                categorization,
                gl_mapping_type,
                cash_categorization,
                cash_type_mapping,
                owner,
                acquired
            FROM {config.GL_RAW_TABLE}
            WHERE property_name = ?
              AND COALESCE(categorization, '') = ?
            ORDER BY txn_date ASC
            LIMIT 2000;
            """,
            conn,
            params=(TARGET_PROPERTY_NAME, TARGET_CATEGORIZATION),
        )
        _print_df("gl_raw: rows for TARGET_PROPERTY_NAME where categorization == 'Rent'", df_raw_rent, max_rows=60)

        # 3) Compare to gl_agg for the same property + categorization
        df_agg_rent = pd.read_sql_query(
            f"""
            SELECT
                month_start,
                timeframe,
                investor,
                owner,
                property_name,
                property,
                acquired,
                categorization,
                gl_mapping_type,
                cash_categorization,
                cash_type_mapping,
                value,
                cash_value
            FROM {config.GL_AGG_TABLE}
            WHERE property_name = ?
              AND COALESCE(categorization, '') = ?
            ORDER BY month_start ASC, investor ASC
            LIMIT 2000;
            """,
            conn,
            params=(TARGET_PROPERTY_NAME, TARGET_CATEGORIZATION),
        )
        _print_df("gl_agg: rows for TARGET_PROPERTY_NAME where categorization == 'Rent'", df_agg_rent, max_rows=120)

        # 4) Quick totals comparison (raw base vs agg sum across investors)
        if not df_raw_rent.empty:
            raw_net = float(df_raw_rent["net"].sum())
        else:
            raw_net = 0.0

        if not df_agg_rent.empty:
            agg_value_sum = float(df_agg_rent["value"].sum())
        else:
            agg_value_sum = 0.0

        print("\n" + "=" * 90)
        print("Totals check")
        print("=" * 90)
        print(f"Raw Rent net (SUM(debit-credit)) for property: {raw_net:,.2f}")
        print(f"Agg Rent value sum (SUM(value)) across investors: {agg_value_sum:,.2f}")

        # 5) Check if owner/acquired are null in raw (would break allocation join)
        df_nulls = pd.read_sql_query(
            f"""
            SELECT
                SUM(CASE WHEN owner IS NULL OR owner = '' THEN 1 ELSE 0 END) AS owner_null_rows,
                SUM(CASE WHEN acquired IS NULL OR acquired = '' THEN 1 ELSE 0 END) AS acquired_null_rows,
                COUNT(*) AS total_rows
            FROM {config.GL_RAW_TABLE}
            WHERE property_name = ?;
            """,
            conn,
            params=(TARGET_PROPERTY_NAME,),
        )
        _print_df("gl_raw: null check for owner/acquired for TARGET_PROPERTY_NAME", df_nulls, max_rows=10)

    finally:
        conn.close()


if __name__ == "__main__":
    main()