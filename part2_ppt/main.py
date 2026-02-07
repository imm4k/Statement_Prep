from __future__ import annotations

from pathlib import Path

from pptx import Presentation

import config
from excel_inputs import load_general_config, load_investors_from_run_config
from ppt_append import combine_presentations
from ppt_objects import UpdateContext, apply_object_updates

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

    investors = load_investors_from_run_config(
        xlsx_path=setup_xlsx,
        sheet_name=config.RUN_CONFIG_SHEET,
    )

    statement_thru_yyyymm = general.statement_thru_date.strftime("%Y_%m")
    standard_slides_path = config.TEMPLATE_DIR / config.STANDARD_SLIDES_FILENAME
    if not standard_slides_path.exists():
        raise FileNotFoundError(f"Missing standard slides deck: {standard_slides_path}")

    print(f"Statement Thru Date: {general.statement_thru_date.strftime('%Y-%m-%d')}")
    print(f"Output Location: {general.output_location}")
    print(f"Template Dir: {config.TEMPLATE_DIR}")
    print(f"Investors to process: {len(investors)}")

    for idx, investor in enumerate(investors, start=1):
        print(f"Starting investor {investor}. Currently on {idx} of {len(investors)}")

        investor_template_name = config.INVESTOR_TEMPLATE_FORMAT.format(investor=investor)
        investor_template_path = config.TEMPLATE_DIR / investor_template_name
        if not investor_template_path.exists():
            raise FileNotFoundError(f"Missing investor template: {investor_template_path}")

        prs = Presentation(str(investor_template_path))

        t1_str = general.statement_thru_date.strftime("%b %Y")

        ctx = UpdateContext(
            investor=investor,
            owner=None,
            statement_thru_date_dt=general.statement_thru_date,
            statement_thru_date_str=general.statement_thru_date.strftime("%m/%d/%Y"),
            t1_str=t1_str,
        )
        apply_object_updates(prs, ctx)

        investor_out_dir = general.output_location / investor
        investor_out_dir.mkdir(parents=True, exist_ok=True)

        out_name = config.DEFAULT_OUTPUT_FILENAME_FORMAT.format(
            statement_thru_yyyymm=statement_thru_yyyymm,
            investor=investor,
        )
        out_path = investor_out_dir / out_name

        tmp_updated_path = investor_out_dir / f"__tmp_updated_{out_name}"

        if tmp_updated_path.exists():
            tmp_updated_path.unlink()

        prs.save(str(tmp_updated_path))
        print(f"Saved temp updated deck: {tmp_updated_path}")

        combine_presentations(
            base_pptx_path=tmp_updated_path,
            standard_pptx_path=standard_slides_path,
            out_pptx_path=out_path,
        )

        tmp_updated_path.unlink()

    print("All investors completed.")


if __name__ == "__main__":
    main()
