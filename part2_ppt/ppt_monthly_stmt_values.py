from __future__ import annotations

import sqlite3
from datetime import date
from typing import Dict, List, Optional, Tuple

import config
from ppt_objects import UpdateContext


_TIMEFRAMES = [f"[T{n}]" for n in range(1, 14)]


def list_investor_owner_property_triplets() -> List[Tuple[str, str, str]]:
    """
    Returns all distinct (investor, owner, property) combinations present in gl_agg.
    Property is treated as the lowest granularity key for the export.
    """
    sql = """
        SELECT DISTINCT
            TRIM(COALESCE(investor, '')) AS investor,
            TRIM(COALESCE(owner, '')) AS owner,
            TRIM(COALESCE(property, '')) AS property
        FROM gl_agg
        WHERE TRIM(COALESCE(investor, '')) <> ''
          AND TRIM(COALESCE(owner, '')) <> ''
          AND TRIM(COALESCE(property, '')) <> ''
          AND (timeframe IS NULL OR timeframe <> 'N/A')
        ORDER BY investor, owner, property
    """

    con = sqlite3.connect(str(config.SQLITE_PATH))
    try:
        rows = con.execute(sql).fetchall()
    finally:
        con.close()

    out: List[Tuple[str, str, str]] = []
    for inv, own, prop in rows:
        out.append((str(inv).strip(), str(own).strip(), str(prop).strip()))
    return out


def _owner_filter_sql(ctx: UpdateContext) -> tuple[str, tuple]:
    if ctx.owner is None or str(ctx.owner).strip() == "":
        return "", tuple()
    return " AND owner = ? ", (str(ctx.owner).strip(),)


def _property_filter_sql(property_name: Optional[str]) -> tuple[str, tuple]:
    if property_name is None or str(property_name).strip() == "":
        return "", tuple()
    return " AND property = ? ", (str(property_name).strip(),)


def _timeframe_list_sql() -> str:
    return ",".join([f"'{tf}'" for tf in _TIMEFRAMES])


def build_month_year_labels(ctx: UpdateContext, property_name: Optional[str]) -> Dict[str, str]:
    """
    Returns mapping: [Tn] -> "Mon YYYY"
    Matches the PPT logic: uses MAX(month_start) for each timeframe token.
    """
    owner_sql, owner_params = _owner_filter_sql(ctx)
    prop_sql, prop_params = _property_filter_sql(property_name)

    sql = f"""
        SELECT timeframe, MAX(month_start) AS month_start
        FROM gl_agg
        WHERE investor = ?
          AND timeframe IS NOT NULL
          AND timeframe <> 'N/A'
          AND timeframe IN ({_timeframe_list_sql()})
          AND month_start IS NOT NULL
          {owner_sql}
          {prop_sql}
        GROUP BY timeframe
    """

    con = sqlite3.connect(str(config.SQLITE_PATH))
    try:
        rows = con.execute(sql, (ctx.investor, *owner_params, *prop_params)).fetchall()
    finally:
        con.close()

    token_to_label: Dict[str, str] = {}
    for tf, ms in rows:
        try:
            y = int(str(ms)[:4])
            m = int(str(ms)[5:7])
            dt = date(y, m, 1)
            token_to_label[str(tf).strip()] = dt.strftime("%b %Y")
        except Exception:
            continue

    return token_to_label


def build_monthly_perf_totals(ctx: UpdateContext, property_name: Optional[str]) -> Dict[str, float]:
    """
    Returns totals for the full T1 to T13 aggregation, per property if property_name provided.
    Keys match the PPT columns.
    """
    owner_sql, owner_params = _owner_filter_sql(ctx)
    prop_sql, prop_params = _property_filter_sql(property_name)

    wanted_cats = (
        "Rent",
        "Dividend",
        "HOA & Mgt. Fee",
        "Repairs & Other Exp.",
        "Mortgage Interest",
    )
    placeholders = ",".join(["?"] * len(wanted_cats))

    sql = f"""
        SELECT categorization, gl_mapping_type, SUM(value) AS total_value
        FROM gl_agg
        WHERE investor = ?
          AND (timeframe IS NULL OR timeframe <> 'N/A')
          AND timeframe IN ({_timeframe_list_sql()})
          AND categorization IN ({placeholders})
          {owner_sql}
          {prop_sql}
        GROUP BY categorization, gl_mapping_type
    """

    con = sqlite3.connect(str(config.SQLITE_PATH))
    try:
        rows = con.execute(sql, (ctx.investor, *wanted_cats, *owner_params, *prop_params)).fetchall()
    finally:
        con.close()

    cat_totals: Dict[str, float] = {k: 0.0 for k in wanted_cats}

    for cat, mapping_type, total_value in rows:
        cat_key = str(cat).strip()
        mt = str(mapping_type or "").strip().lower()
        v = float(total_value or 0.0)

        if mt in ("revenue", "expense"):
            v = -1.0 * v

        if cat_key in cat_totals:
            cat_totals[cat_key] += v

    rent = float(cat_totals.get("Rent", 0.0))
    dividend = float(cat_totals.get("Dividend", 0.0))
    hoa_mgt = float(cat_totals.get("HOA & Mgt. Fee", 0.0))
    repairs_other = float(cat_totals.get("Repairs & Other Exp.", 0.0))
    mortgage_int = float(cat_totals.get("Mortgage Interest", 0.0))

    total_rev = rent + dividend
    total_exp = hoa_mgt + repairs_other + mortgage_int
    monthly = total_rev + total_exp
    cumulative = monthly

    return {
        "Rent": rent,
        "Dividend": dividend,
        "Total Revenue": total_rev,
        "HOA & Mgt. Fee": hoa_mgt,
        "Repairs & Other Exp.": repairs_other,
        "Mortgage Interest": mortgage_int,
        "Total Expenses": total_exp,
        "Monthly": monthly,
        "Cumulative": cumulative,
    }


def build_monthly_cash_totals(ctx: UpdateContext, property_name: Optional[str]) -> Dict[str, float]:
    """
    Returns totals for the full T1 to T13 aggregation, per property if property_name provided.
    Keys match the PPT columns.
    """
    owner_sql, owner_params = _owner_filter_sql(ctx)
    prop_sql, prop_params = _property_filter_sql(property_name)

    sql = f"""
        SELECT
            cash_categorization,
            cash_type_mapping,
            SUM(cash_value) AS total_value
        FROM gl_agg
        WHERE investor = ?
          AND (timeframe IS NULL OR timeframe <> 'N/A')
          AND timeframe IN ({_timeframe_list_sql()})
          {owner_sql}
          {prop_sql}
        GROUP BY cash_categorization, cash_type_mapping
    """

    con = sqlite3.connect(str(config.SQLITE_PATH))
    try:
        rows = con.execute(sql, (ctx.investor, *owner_params, *prop_params)).fetchall()
    finally:
        con.close()

    by_cat: Dict[str, float] = {}
    inflow_total = 0.0
    outflow_total = 0.0

    for cat, cash_type, total_value in rows:
        cat_key = str(cat or "").strip()
        ct = str(cash_type or "").strip().lower()
        v_raw = -1.0 * float(total_value or 0.0)
        v_abs = abs(v_raw)

        # Special handling: Mortgage Principal can be "Both" and must land in either
        # "Mortgage Loan" (if positive) or "Mortgage Principal" (if negative).
        if ct == "both" and cat_key == "Mortgage Principal":
            if v_raw > 0:
                cat_key = "Mortgage Loan"
                v_signed = v_abs
                inflow_total += v_abs
            elif v_raw < 0:
                cat_key = "Mortgage Principal"
                v_signed = -1.0 * v_abs
                outflow_total += (-1.0 * v_abs)
            else:
                v_signed = 0.0
        else:
            v_signed = 0.0
            if ct == "inflow":
                v_signed = v_abs
                inflow_total += v_abs
            elif ct == "outflow":
                v_signed = -1.0 * v_abs
                outflow_total += (-1.0 * v_abs)
            else:
                v_signed = 0.0

        if cat_key != "":
            by_cat[cat_key] = by_cat.get(cat_key, 0.0) + v_signed

    owner_contrib = float(by_cat.get("Owner Contribution", 0.0))
    mortgage_loan = float(by_cat.get("Mortgage Loan", 0.0))
    rent_dividend = float(by_cat.get("Rent & Dividend", 0.0))

    hoa_mgt = float(by_cat.get("HOA & Mgt. Fee", 0.0))
    repairs_other = float(by_cat.get("Repairs & Other Exp.", 0.0))
    mortgage_interest = float(by_cat.get("Mortgage Interest", 0.0))
    mortgage_principal = float(by_cat.get("Mortgage Principal", by_cat.get("Mortgage Principle", 0.0)))
    apartment_improve = float(by_cat.get("Apartment & Improve.", 0.0))
    owner_distribution = float(by_cat.get("Owner Distribution", 0.0))

    total_inflow = float(inflow_total)
    total_outflow = float(outflow_total)
    monthly = total_inflow + total_outflow
    cumulative = monthly

    return {
        "Owner Contribution": owner_contrib,
        "Mortgage Loan": mortgage_loan,
        "Rent & Dividend": rent_dividend,
        "Total Inflow": total_inflow,
        "HOA & Mgt. Fee": hoa_mgt,
        "Repairs & Other Exp.": repairs_other,
        "Mortgage Interest": mortgage_interest,
        "Mortgage Principal": mortgage_principal,
        "Apartment & Improve.": apartment_improve,
        "Owner Distribution": owner_distribution,
        "Total Outflow": total_outflow,
        "Monthly": monthly,
        "Cumulative": cumulative,
    }
