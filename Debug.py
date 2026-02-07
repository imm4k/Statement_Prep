import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
from common.sqlite_utils import connect

SETUP_XLSX_PATH = r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep_setup.xlsx"
SQLITE_PATH = r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep.sqlite"
GL_RAW_TABLE = "gl_raw"


def main() -> None:
    print("Loading Investor Table from Excel...")
    inv = pd.read_excel(SETUP_XLSX_PATH, sheet_name="Investor Table", engine="openpyxl")

    required = ["Investor", "Property Name", "Owner", "Acquired", "Type"]
    missing = [c for c in required if c not in inv.columns]
    if missing:
        print("Investor Table missing columns:", missing)
        print("Investor Table columns:", inv.columns.tolist())
        return

    inv = inv[required].copy()
    inv["Property Name"] = inv["Property Name"].astype(str).str.strip()

    print("Investor Table rows:", len(inv))
    print("Investor Table sample Property Name (first 5):")
    for i, v in enumerate(inv["Property Name"].head(5).tolist(), start=1):
        print(f"{i}: {repr(v)}")

    print("\nLoading distinct property_name values from SQLite gl_raw...")
    conn = connect(SQLITE_PATH)
    try:
        gl_props = pd.read_sql_query(
            f"SELECT DISTINCT property_name FROM {GL_RAW_TABLE} WHERE property_name IS NOT NULL;",
            conn,
        )
        mapped_preview = pd.read_sql_query(
            f"""
            SELECT
                property_name,
                investor,
                owner
            FROM {GL_RAW_TABLE}
            WHERE property_name IS NOT NULL
            LIMIT 10;
            """,
            conn,
        )
        null_counts = pd.read_sql_query(
            f"""
            SELECT
                SUM(CASE WHEN investor IS NULL OR TRIM(investor) = '' THEN 1 ELSE 0 END) AS investor_null_rows,
                SUM(CASE WHEN owner IS NULL OR TRIM(owner) = '' THEN 1 ELSE 0 END) AS owner_null_rows,
                COUNT(*) AS total_rows
            FROM {GL_RAW_TABLE};
            """,
            conn,
        )
    finally:
        conn.close()

    gl_props["property_name"] = gl_props["property_name"].astype(str).str.strip()

    print("Distinct gl_raw property_name count:", len(gl_props))
    print("gl_raw sample property_name (first 5):")
    for i, v in enumerate(gl_props["property_name"].head(5).tolist(), start=1):
        print(f"{i}: {repr(v)}")

    print("\nCurrent mapping status in gl_raw (first 10 rows):")
    print(mapped_preview.to_string(index=False))

    print("\nNull counts in gl_raw:")
    print(null_counts.to_string(index=False))

    print("\nChecking match rate: gl_raw.property_name -> Investor Table.Property Name ...")
    inv_set = set(inv["Property Name"].tolist())
    gl_set = set(gl_props["property_name"].tolist())

    matched = len(gl_set.intersection(inv_set))
    print("Distinct gl_raw property_name values:", len(gl_set))
    print("Distinct Investor Table Property Name values:", len(inv_set))
    print("Distinct property_name values that match:", matched)

    unmatched = sorted(list(gl_set - inv_set))
    print("\nUnmatched gl_raw property_name (first 25):")
    for i, v in enumerate(unmatched[:25], start=1):
        print(f"{i}: {repr(v)}")

    print("\nIf unmatched list is non-empty, the join cannot populate investor/owner.")
    print("If unmatched list is empty but investor/owner still null, the UPDATE query is not executing or is targeting a different table.")

if __name__ == "__main__":
    main()
