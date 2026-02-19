from __future__ import annotations

from typing import Dict

from ppt_objects import ObjectUpdater

from ppt_object_logic_text import (
    update_cover_title,
    update_cover_subtitle,
    update_summary_title,
    update_cash_summary_title,
    update_summary_top_text,
)

from ppt_object_logic_tables import (
    update_summary_table,
    update_monthly_perf_table,
    update_monthly_cash_table,
    update_available_cash,
)

from ppt_object_logic_visibility import update_partial_ownership_visibility

OBJECT_UPDATERS: Dict[str, ObjectUpdater] = {
    "cover_title": update_cover_title,
    "cover_subtitle": update_cover_subtitle,
    "summary_title": update_summary_title,
    "summary_table": update_summary_table,
    "summary_top_text": update_summary_top_text,
    "cash_summary_title": update_cash_summary_title,
    "available_cash": update_available_cash,
    "monthly_cash_table": update_monthly_cash_table,
    "monthly_perf_table": update_monthly_perf_table,

    "pct_owner_note": update_partial_ownership_visibility,
    "overview_title": update_partial_ownership_visibility,
    "pct_overview_title": update_partial_ownership_visibility,
}
