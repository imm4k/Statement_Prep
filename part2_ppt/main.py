from __future__ import annotations

import sys
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent               # ...\Statement_Prep\part2_ppt
_ROOT_DIR = _THIS_DIR.parent                              # ...\Statement_Prep
_COMMON_DIR = _ROOT_DIR / "common"                        # ...\Statement_Prep\common

for p in (str(_COMMON_DIR), str(_ROOT_DIR), str(_THIS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

from pptx import Presentation

import config
from excel_inputs import load_general_config, load_run_config_rows, load_investor_table_ownership_map
from ppt_append import combine_presentations
from ppt_objects import UpdateContext, apply_object_updates

def _sanitize_filename_component(s: str) -> str:
    bad = '<>:"/\\|?*'
    out = "".join("_" if ch in bad else ch for ch in str(s or "").strip())
    out = out.replace("  ", " ").strip()
    return out

def main() -> None:
    setup_xlsx = Path(config.SETUP_EXCEL_PATH)
    if not setup_xlsx.exists():
        raise FileNotFoundError(f"Missing setup workbook: {setup_xlsx}")

    general = load_general_config(
        xlsx_path=setup_xlsx,
        sheet_name=config.GENERAL_CONFIG_SHEET,
        label_output_location=config.GENERAL_CONFIG_LABEL_OUTPUT_LOCATION,
        label_statement_thru=config.GENERAL_CONFIG_LABEL_STATEMENT_THRU_DATE,
    )

    run_rows = load_run_config_rows(
        xlsx_path=setup_xlsx,
        sheet_name=config.RUN_CONFIG_SHEET,
    )

    ownership_map = load_investor_table_ownership_map(setup_xlsx)

    statement_thru_yyyymm = general.statement_thru_date.strftime("%Y_%m")
    standard_slides_path = config.TEMPLATE_DIR / config.STANDARD_SLIDES_FILENAME
    if not standard_slides_path.exists():
        raise FileNotFoundError(f"Missing standard slides deck: {standard_slides_path}")

    print(f"Statement Thru Date: {general.statement_thru_date.strftime('%Y-%m-%d')}")
    print(f"Output Location: {general.output_location}")
    print(f"Template Dir: {config.TEMPLATE_DIR}")
    print(f"Runs to process: {len(run_rows)}")

    completed = 0

    for idx, run in enumerate(run_rows, start=1):

        inv_key = str(run.investor or "").strip()
        owner_key = str(run.owner or "").strip()
        if inv_key == "":
            raise ValueError(f"Run row {idx} missing Investor value.")
        if owner_key == "":
            raise ValueError(f"Run row {idx} missing Owner value for investor '{inv_key}'.")

        map_key = (inv_key.lower(), owner_key.lower())
        if map_key not in ownership_map:
            raise ValueError(
                f"Could not find any properties for Investor Owner in Investor Table for run. "
                f"Investor: {inv_key}. Owner: {owner_key}"
            )

        pcts = ownership_map[map_key]
        if not pcts:
            raise ValueError(
                f"Investor Table returned no % Ownership values for run. "
                f"Investor: {inv_key}. Owner: {owner_key}"
            )

        all_full = all(float(x) >= 100.0 for x in pcts)
        ownership_pct = 100.0 if all_full else float(min(pcts))
        ownership_factor = ownership_pct / 100.0

        print(
            f"Starting run {idx} of {len(run_rows)}. Investor: {inv_key}. Owner: {owner_key}. % Ownership (Investor Table): {ownership_pct}%"
        )

        owner_template_name = config.OWNER_TEMPLATE_FORMAT.format(owner=_sanitize_filename_component(owner_key))
        owner_template_path = config.TEMPLATE_DIR / owner_template_name
        if not owner_template_path.exists():
            raise FileNotFoundError(f"Missing owner template: {owner_template_path}")

        prs = Presentation(str(owner_template_path))

        t1_str = general.statement_thru_date.strftime("%b %Y")

        ctx = UpdateContext(
            investor=inv_key,
            owner=owner_key,
            ownership_pct=ownership_pct,
            ownership_factor=ownership_factor,
            statement_thru_date_dt=general.statement_thru_date,
            statement_thru_date_str=general.statement_thru_date.strftime("%m/%d/%Y"),
            t1_str=t1_str,
        )

        apply_object_updates(prs, ctx)

        investor_out_dir = general.output_location / inv_key
        investor_out_dir.mkdir(parents=True, exist_ok=True)

        owner_for_filename = _sanitize_filename_component(owner_key)

        out_name = config.DEFAULT_OUTPUT_FILENAME_FORMAT.format(
            statement_thru_yyyymm=statement_thru_yyyymm,
            owner=owner_for_filename,
        )
        out_path = investor_out_dir / out_name

        tmp_updated_path = investor_out_dir / f"__tmp_updated_{out_name}"
        if tmp_updated_path.exists():
            tmp_updated_path.unlink()

        prs.save(str(tmp_updated_path))
        print(f"Saved temp updated deck: {tmp_updated_path}")

        chosen_standard = (run.standard_slides_template or "").strip()
        if chosen_standard == "":
            chosen_standard = config.STANDARD_SLIDES_FILENAME

        chosen_standard_path = config.TEMPLATE_DIR / chosen_standard
        if not chosen_standard_path.exists():
            print(f"Standard slides template not found. Falling back to default. Missing: {chosen_standard_path}")
            chosen_standard_path = config.TEMPLATE_DIR / config.STANDARD_SLIDES_FILENAME

        combine_presentations(
            base_pptx_path=str(tmp_updated_path),
            standard_pptx_path=str(chosen_standard_path),
            out_pptx_path=str(out_path),
            ownership_pct=float(ownership_pct),
        )

        tmp_updated_path.unlink()
        completed += 1

    if bool(getattr(config, "EXPORT_MONTHLY_STMT_XLSX", False)):
        part2_dir = Path(__file__).resolve().parent
        root_dir = part2_dir.parent
        common_dir = root_dir / "common"

        if str(common_dir) not in sys.path:
            sys.path.insert(0, str(common_dir))

        from monthly_stmt_export import export_monthly_stmt_excel

        export_monthly_stmt_excel()

    print(f"All investors completed. Successful runs: {completed} of {len(run_rows)}")

if __name__ == "__main__":
    main()
