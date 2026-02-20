from __future__ import annotations

from pathlib import Path

import pandas as pd


GL_CSV_PATH = Path(
    r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\23. Operations and Administrative\Property Management\Investor Updates\2026_01\general_ledger-20260218.csv"
)

SETUP_XLSX_PATH = Path(
    r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep_setup.xlsx"
)

GL_MAPPING_SHEET = "GL Mapping"

PROPERTY_NAME_FILTER = "CPM Luca/Jamie/David 2- Elle - 1558upr5a"

OUTPUT_DIR = Path(
    r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\23. Operations and Administrative\Property Management\Investor Updates\2026_01\David Liang"
)

OUTPUT_XLSX_NAME = "debug_gl_records_with_mapping.xlsx"


def main() -> None:
    if not GL_CSV_PATH.exists():
        raise FileNotFoundError(f"Missing GL CSV: {GL_CSV_PATH}")
    if not SETUP_XLSX_PATH.exists():
        raise FileNotFoundError(f"Missing setup workbook: {SETUP_XLSX_PATH}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUTPUT_DIR / OUTPUT_XLSX_NAME

    print("=================================================================")
    print("debug.py: GL record filter + GL Mapping categorization join")
    print("=================================================================")
    print(f"GL CSV:     {GL_CSV_PATH}")
    print(f"Setup XLSX: {SETUP_XLSX_PATH} (sheet '{GL_MAPPING_SHEET}')")
    print(f"Filter:     Property Name == '{PROPERTY_NAME_FILTER}'")
    print(f"Output:     {out_path}")
    print()

    # User instruction: CSV row 1 is headers, actual records start on row 4
    # Therefore skip rows 2 and 3 (1-indexed), which are indices 1 and 2 (0-indexed) after the header line.
    print("Loading GL CSV (skipping rows 2 and 3 after header)...")
    gl = pd.read_csv(
        GL_CSV_PATH,
        dtype=str,
        header=0,
        skiprows=[1, 2],
        keep_default_na=False,
        na_values=[],
    )
    print(f"Loaded GL rows: {len(gl):,}")
    print(f"GL columns: {list(gl.columns)}")
    print()

    required_gl_cols = ["Property Name", "GL Account"]
    for col in required_gl_cols:
        if col not in gl.columns:
            raise KeyError(f"GL CSV missing required column: '{col}'")

    print("Filtering GL rows by Property Name...")
    gl_f = gl[gl["Property Name"].astype(str).str.strip() == PROPERTY_NAME_FILTER].copy()
    print(f"Filtered rows: {len(gl_f):,}")
    print()

    print("Loading GL Mapping sheet...")
    mapping = pd.read_excel(
        SETUP_XLSX_PATH,
        sheet_name=GL_MAPPING_SHEET,
        dtype=str,
        keep_default_na=False,
        na_values=[],
    )
    print(f"Loaded GL Mapping rows: {len(mapping):,}")
    print(f"GL Mapping columns: {list(mapping.columns)}")
    print()

    required_map_cols = ["GL Account", "Categorization", "Cash Categorization"]
    for col in required_map_cols:
        if col not in mapping.columns:
            raise KeyError(f"GL Mapping sheet missing required column: '{col}'")

    mapping_slim = mapping[["GL Account", "Categorization", "Cash Categorization"]].copy()

    print("Joining GL rows to GL Mapping on 'GL Account'...")
    gl_f["GL Account"] = gl_f["GL Account"].astype(str).str.strip()
    mapping_slim["GL Account"] = mapping_slim["GL Account"].astype(str).str.strip()

    out = gl_f.merge(mapping_slim, on="GL Account", how="left")

    # Move new columns to the end (as requested: add columns)
    cols = [c for c in out.columns if c not in ("Categorization", "Cash Categorization")]
    cols += ["Categorization", "Cash Categorization"]
    out = out[cols]

    print("Writing Excel output...")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="gl_filtered", index=False)

    print()
    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    main()