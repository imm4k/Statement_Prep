from __future__ import annotations

from typing import Dict

from ppt_objects import ObjectUpdater

from ppt_object_logic_text import (
    update_cover_title,
    update_cover_subtitle,

    update_overview_title,
    update_overview_title_pct,

    update_perf_summary_title,
    update_perf_summary_title_pct,

    update_cash_summary_title,
    update_cash_summary_title_pct,

    update_summary_top_text,
)

from ppt_object_logic_tables import (
    update_summary_table,
    update_nav_table,
    update_ni_table,
    update_ca_table,
    update_monthly_perf_table,
    update_monthly_cash_table,
    update_available_cash,
)

from ppt_object_logic_visibility import update_partial_ownership_visibility


def _chain(*updaters: ObjectUpdater) -> ObjectUpdater:
    def _runner(slide, shape, prs, ctx) -> None:
        for u in updaters:
            u(slide, shape, prs, ctx)
    return _runner


OBJECT_UPDATERS: Dict[str, ObjectUpdater] = {
    "cover_title": update_cover_title,
    "cover_subtitle": update_cover_subtitle,

    # Titles: always update text, then enforce full vs partial visibility
    "overview_title": _chain(update_overview_title, update_partial_ownership_visibility),
    "overview_title_pct": _chain(update_overview_title_pct, update_partial_ownership_visibility),

    "perf_summary_title": _chain(update_perf_summary_title, update_partial_ownership_visibility),
    "perf_summary_title_pct": _chain(update_perf_summary_title_pct, update_partial_ownership_visibility),

    "cash_summary_title": _chain(update_cash_summary_title, update_partial_ownership_visibility),
    "cash_summary_title_pct": _chain(update_cash_summary_title_pct, update_partial_ownership_visibility),

    # Tables and other objects
    "summary_table": update_summary_table,
    "nav_table": update_nav_table,
    "ni_table": update_ni_table,
    "ca_table": update_ca_table,
    "summary_top_text": update_summary_top_text,

    "available_cash": update_available_cash,
    "monthly_cash_table": update_monthly_cash_table,
    "monthly_perf_table": update_monthly_perf_table,

    # Visibility only
    "pct_owner_note": update_partial_ownership_visibility,
}