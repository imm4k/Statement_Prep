from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import os
from datetime import datetime

import pandas as pd

from common.sqlite_utils import connect

DEFAULT_DB_PATH = r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep.sqlite"
DEFAULT_TABLE_NAME = "gl_agg"
DEFAULT_OUTPUT_DIR = r"C:\Users\liang\Downloads"


def export_gl_agg_to_excel(
    db_path: str = DEFAULT_DB_PATH,
    table_name: str = DEFAULT_TABLE_NAME,
    output_path: str | None = None,
) -> str:
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(DEFAULT_OUTPUT_DIR, f"{table_name}_export_{ts}.xlsx")

    conn = connect(db_path)
    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name};", conn)
    finally:
        conn.close()

    df.to_excel(output_path, index=False)
    print(f"Exported {len(df)} rows to: {output_path}")
    return output_path

if __name__ == "__main__":
    # Manually set output_path here when you want a specific filename.
    export_gl_agg_to_excel()

