from __future__ import annotations

from pathlib import Path


TEMPLATE_DIR = Path(
    r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\23. Operations and Administrative\Property Management\Investor Updates\0. Templates"
)

STANDARD_SLIDES_FILENAME = "0. TEMPLATE_Monthly_Standard_Slides.pptx"
OWNER_TEMPLATE_FORMAT = "TEMPLATE_Monthly_{owner}.pptx"

SETUP_EXCEL_PATH = Path(
    r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep_setup.xlsx"
)

SQLITE_PATH = Path(r"H:\.shortcut-targets-by-id\1Tf1JC85Wg2bbW79jfilWWqexA8jYAOxs\CARTER Property Management\0. Company Assets\Automations\Statement Prep\statement_prep.sqlite")


GENERAL_CONFIG_SHEET = "General Config"
RUN_CONFIG_SHEET = "Run Config"

GENERAL_CONFIG_LABEL_OUTPUT_LOCATION = "Output Location:"
GENERAL_CONFIG_LABEL_STATEMENT_THRU_DATE = "Statement Thru Date:"

DEFAULT_OUTPUT_FILENAME_FORMAT = "Monthly_{statement_thru_yyyymm}_{owner}.pptx"