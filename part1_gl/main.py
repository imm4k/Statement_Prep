from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import os

from common.excel_config import load_setup_config
from part1_gl.config import (
    GL_AGG_TABLE,
    GL_RAW_TABLE,
    ROWS_TO_SKIP_AFTER_HEADER,
    SETUP_XLSX_PATH,
    SQLITE_PATH,
)
from part1_gl.gl_enrich_and_aggregate import apply_mappings_inplace, build_aggregate_table
from part1_gl.gl_ingest import ingest_gl_csv_to_raw


def _create_investor_folders(output_location: str, investors: list[str]) -> dict[str, str]:
    total = len(investors)
    print(f"Starting process for {total} investors.")
    out = {}

    for idx, investor in enumerate(investors, start=1):
        print(f"Currently on {idx} of {total}. Investor: {investor}")
        folder = os.path.join(output_location, investor)
        os.makedirs(folder, exist_ok=True)
        out[investor] = folder

    return out


def run_part1() -> None:
    setup = load_setup_config(SETUP_XLSX_PATH)

    investor_folders = _create_investor_folders(setup.general.output_location, setup.investors)

    csv_path = os.path.join(setup.general.gl_location, setup.general.gl_file_name)

    print("Starting GL ingest into SQLite.")
    ingest_gl_csv_to_raw(
        db_path=SQLITE_PATH,
        gl_raw_table=GL_RAW_TABLE,
        csv_path=csv_path,
        rows_to_skip_after_header=ROWS_TO_SKIP_AFTER_HEADER,
    )
    print("Completed GL ingest into SQLite.")

    print("Starting enrichment mappings.")
    apply_mappings_inplace(
        db_path=SQLITE_PATH,
        gl_raw_table=GL_RAW_TABLE,
        investor_table_df=setup.investor_table,
        gl_mapping_df=setup.gl_mapping,
    )
    print("Completed enrichment mappings.")

    print("Starting aggregation build.")
    build_aggregate_table(
        db_path=SQLITE_PATH,
        gl_raw_table=GL_RAW_TABLE,
        gl_agg_table=GL_AGG_TABLE,
    )
    print("Completed aggregation build.")

    print("Part 1 completed.")


if __name__ == "__main__":
    run_part1()
